// Motor Estadístico Completo — Estadística I
// Soporta carga de .xlsx y .csv, detección de tipos de variable, validación de datos,
// tablas de frecuencias simples y agrupadas, medidas completas, gráficos e informe.

(function () {
  "use strict";

  const state = {
    fileName: "encuesta_ia_educacion_respuestas.xlsx (Precargado - 75 respuestas)",
    headers: [],
    rawRows: [],
    cleanRows: [],
    variables: [], // { index, name, autoType, userType, chartType, validCount, emptyCount, invalidCount, uniqueValues }
    chartInstances: {},
    selectedQuestionFilter: "all",
    encodeYesNoAsNumeric: false
  };

  const PALETTE = [
    "#3b4cca", "#0ea5e9", "#10b981", "#f59e0b",
    "#8b5cf6", "#ec4899", "#14b8a6", "#f97316",
    "#6366f1", "#84cc16"
  ];

  // ==========================================
  // 1. UTILIDADES MATEMÁTICAS Y ESTADÍSTICAS
  // ==========================================

  function round(val, decimals = 4) {
    if (val === null || val === undefined || Number.isNaN(val)) return 0;
    const factor = Math.pow(10, decimals);
    return Math.round((val + Number.EPSILON) * factor) / factor;
  }

  function parseNumericValue(val, encodeYesNo = false) {
    if (val === null || val === undefined) return NaN;
    if (typeof val === "number") return Number.isFinite(val) ? val : NaN;
    const str = String(val).trim();
    if (str === "") return NaN;

    if (encodeYesNo) {
      const lower = str.toLowerCase();
      if (lower === "sí" || lower === "si" || lower === "yes") return 1;
      if (lower === "no") return 0;
    }

    const normalized = str.replace(",", ".");
    if (/^-?\d+(\.\d+)?$/.test(normalized)) {
      return parseFloat(normalized);
    }
    return NaN;
  }

  // Cálculo de percentil por interpolación lineal (Método estándar R-7 / Excel)
  function calcPercentile(sortedArr, p) {
    const n = sortedArr.length;
    if (n === 0) return 0;
    if (n === 1) return sortedArr[0];
    const pos = (p / 100) * (n - 1);
    const base = Math.floor(pos);
    const rest = pos - base;
    if (sortedArr[base + 1] !== undefined) {
      return sortedArr[base] + rest * (sortedArr[base + 1] - sortedArr[base]);
    }
    return sortedArr[base];
  }

  // Cálculo completo para variables cuantitativas (Sección 10.2)
  function analyzeQuantitative(values, forceGrouped = false) {
    const n = values.length;
    const sorted = [...values].sort((a, b) => a - b);

    // B. Medidas de tendencia central
    const sum = sorted.reduce((acc, v) => acc + v, 0);
    const mean = n > 0 ? sum / n : 0;
    const median = calcPercentile(sorted, 50);

    // Moda(s)
    const counts = new Map();
    let maxFreq = 0;
    for (const v of sorted) {
      const c = (counts.get(v) || 0) + 1;
      counts.set(v, c);
      if (c > maxFreq) maxFreq = c;
    }
    const modes = [];
    if (maxFreq > 1) {
      for (const [k, c] of counts.entries()) {
        if (c === maxFreq) modes.push(k);
      }
    }
    const modeText = modes.length === 0
      ? "Amodal (sin repeticiones)"
      : `${modes.join(", ")} (f = ${maxFreq})`;

    // D. Medidas de posición
    const min = sorted[0] ?? 0;
    const max = sorted[n - 1] ?? 0;
    const q1 = calcPercentile(sorted, 25);
    const q2 = median;
    const q3 = calcPercentile(sorted, 75);
    const p10 = calcPercentile(sorted, 10);
    const p25 = q1;
    const p50 = median;
    const p75 = q3;
    const p90 = calcPercentile(sorted, 90);

    // C. Medidas de variabilidad
    const range = max - min;
    let sumSqDiff = 0;
    for (const v of sorted) {
      sumSqDiff += Math.pow(v - mean, 2);
    }
    const variance = n > 1 ? sumSqDiff / (n - 1) : 0; // Varianza muestral s^2
    const stdDev = Math.sqrt(variance);
    const cv = mean !== 0 ? (stdDev / Math.abs(mean)) * 100 : 0;
    const iqr = q3 - q1; // Rango intercuartílico (RIC)

    // E. Medidas de forma (Asimetría de Fisher y Curtosis de Fisher)
    let skewness = 0;
    let kurtosisExcess = 0;
    if (n > 2 && stdDev > 0) {
      let m3 = 0;
      for (const v of sorted) {
        m3 += Math.pow((v - mean) / stdDev, 3);
      }
      skewness = (n / ((n - 1) * (n - 2))) * m3;
    }
    if (n > 3 && stdDev > 0) {
      let m4 = 0;
      for (const v of sorted) {
        m4 += Math.pow((v - mean) / stdDev, 4);
      }
      const term1 = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3)) * m4;
      const term2 = (3 * Math.pow(n - 1, 2)) / ((n - 2) * (n - 3));
      kurtosisExcess = term1 - term2;
    }
    const kurtosisClassic = kurtosisExcess + 3;

    // Clasificación según asimetría
    let skewnessClass = "Aproximadamente simétrica";
    if (skewness > 0.25) skewnessClass = "Asimétrica positiva (sesgo a la derecha)";
    else if (skewness < -0.25) skewnessClass = "Asimétrica negativa (sesgo a la izquierda)";

    // Clasificación según curtosis
    let kurtosisClass = "Mesocúrtica";
    if (kurtosisExcess > 0.25) kurtosisClass = "Leptocúrtica (más apuntada / concentrada)";
    else if (kurtosisExcess < -0.25) kurtosisClass = "Platicúrtica (más aplanada / dispersa)";

    // Valores atípicos (Regla de Tukey 1.5 * RIC)
    const lowerFence = q1 - 1.5 * iqr;
    const upperFence = q3 + 1.5 * iqr;
    const outliers = sorted.filter(v => v < lowerFence || v > upperFence);

    // A. Distribución de frecuencias (Simple vs Agrupada por Intervalos con Regla de Sturges)
    const uniqueSorted = Array.from(counts.keys()).sort((a, b) => a - b);
    const useGrouped = forceGrouped || uniqueSorted.length > 10;

    let freqRows = [];
    let kClasses = 0;
    let amplitude = 0;

    if (!useGrouped) {
      let cumAbs = 0;
      let cumRel = 0;
      freqRows = uniqueSorted.map(val => {
        const fi = counts.get(val);
        cumAbs += fi;
        const hi = fi / n;
        cumRel += hi;
        return {
          label: String(val),
          classMark: val,
          fi,
          Fi: cumAbs,
          hi: round(hi, 4),
          Hi: round(cumRel, 4),
          pct: round(hi * 100, 2),
          PctCum: round(cumRel * 100, 2)
        };
      });
    } else {
      // Regla de Sturges: k = round(1 + 3.322 * log10(n))
      kClasses = Math.max(4, Math.round(1 + 3.322 * Math.log10(Math.max(n, 2))));
      amplitude = range > 0 ? Math.ceil((range / kClasses) * 10) / 10 : 1;
      if (Number.isInteger(min) && Number.isInteger(max)) {
        amplitude = Math.max(1, Math.ceil((range + 0.0001) / kClasses));
      }

      let cumAbs = 0;
      let cumRel = 0;
      for (let i = 0; i < kClasses; i++) {
        const li = round(min + i * amplitude, 2);
        const ls = round(li + amplitude, 2);
        const isLast = i === kClasses - 1;
        const fi = sorted.filter(v => isLast ? (v >= li && v <= ls + 0.0001) : (v >= li && v < ls)).length;
        cumAbs += fi;
        const hi = fi / n;
        cumRel += hi;
        const xi = round((li + ls) / 2, 2);
        const intervalLabel = isLast ? `[${li} , ${ls}]` : `[${li} , ${ls})`;
        freqRows.push({
          label: intervalLabel,
          classMark: xi,
          fi,
          Fi: cumAbs,
          hi: round(hi, 4),
          Hi: round(cumRel, 4),
          pct: round(hi * 100, 2),
          PctCum: round(cumRel * 100, 2)
        });
      }
    }

    return {
      n,
      useGrouped,
      kClasses,
      amplitude,
      freqRows,
      mean: round(mean, 2),
      median: round(median, 2),
      modes,
      modeText,
      range: round(range, 2),
      variance: round(variance, 2),
      stdDev: round(stdDev, 2),
      cv: round(cv, 2),
      iqr: round(iqr, 2),
      min: round(min, 2),
      max: round(max, 2),
      q1: round(q1, 2),
      q2: round(q2, 2),
      q3: round(q3, 2),
      p10: round(p10, 2),
      p25: round(p25, 2),
      p50: round(p50, 2),
      p75: round(p75, 2),
      p90: round(p90, 2),
      skewness: round(skewness, 3),
      skewnessClass,
      kurtosisExcess: round(kurtosisExcess, 3),
      kurtosisClassic: round(kurtosisClassic, 3),
      kurtosisClass,
      lowerFence: round(lowerFence, 2),
      upperFence: round(upperFence, 2),
      outliers
    };
  }

  // Cálculo para variables cualitativas (Sección 10.1)
  function analyzeQualitative(values) {
    const n = values.length;
    const counts = new Map();
    for (const v of values) {
      const key = String(v).trim();
      counts.set(key, (counts.get(key) || 0) + 1);
    }

    const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
    let cumAbs = 0;
    let cumRel = 0;

    const freqRows = entries.map(([cat, fi]) => {
      cumAbs += fi;
      const hi = n > 0 ? fi / n : 0;
      cumRel += hi;
      return {
        label: cat,
        fi,
        Fi: cumAbs,
        hi: round(hi, 4),
        Hi: round(cumRel, 4),
        pct: round(hi * 100, 2)
      };
    });

    const topCategory = freqRows[0] || { label: "N/A", fi: 0, pct: 0 };
    const minCategory = freqRows[freqRows.length - 1] || { label: "N/A", fi: 0, pct: 0 };

    return {
      n,
      freqRows,
      topCategory,
      minCategory
    };
  }

  // ==========================================
  // 2. DETECCIÓN Y VALIDACIÓN DE VARIABLES (9.1)
  // ==========================================

  function processDataset(headers, rows, sourceName) {
    state.fileName = sourceName;
    state.headers = headers.map(h => String(h || "").trim());

    // Filtrar filas completamente vacías
    const nonEmptyRows = rows.filter(r =>
      Array.isArray(r) && r.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "")
    );
    state.rawRows = nonEmptyRows;

    // Detectar duplicados exactos (comparando todas las columnas excepto marca temporal si se desea, o fila completa)
    let duplicateCount = 0;
    const seenSignatures = new Set();
    for (const r of nonEmptyRows) {
      const sig = JSON.stringify(r);
      if (seenSignatures.has(sig)) duplicateCount++;
      else seenSignatures.add(sig);
    }
    state.duplicateCount = duplicateCount;
    state.cleanRows = nonEmptyRows;

    // Clasificar cada columna
    state.variables = state.headers.map((headerName, colIdx) => {
      const lowerHeader = headerName.toLowerCase();
      const isTimestamp = colIdx === 0 && (
        lowerHeader.includes("marca temporal") ||
        lowerHeader.includes("timestamp") ||
        lowerHeader.includes("fecha")
      );

      let emptyCount = 0;
      let numericCount = 0;
      let validValues = [];
      const uniqueSet = new Set();

      for (const row of state.cleanRows) {
        const rawCell = row[colIdx];
        const strCell = rawCell !== undefined && rawCell !== null ? String(rawCell).trim() : "";
        if (strCell === "") {
          emptyCount++;
          continue;
        }
        uniqueSet.add(strCell);
        const numVal = parseNumericValue(rawCell, state.encodeYesNoAsNumeric);
        if (!Number.isNaN(numVal)) {
          numericCount++;
        }
        validValues.push(rawCell);
      }

      const validCount = validValues.length;
      const isQuantitative = !isTimestamp && validCount > 0 && (numericCount / validCount) >= 0.8;
      const invalidCount = isQuantitative ? (validCount - numericCount) : 0;

      let autoType = "qualitative";
      if (isTimestamp) {
        autoType = "omit";
      } else if (isQuantitative) {
        // Si tiene más de 10 valores distintos o menciona "horas", sugerir agrupada por intervalos
        if (uniqueSet.size > 10 || lowerHeader.includes("horas")) {
          autoType = "quant_grouped";
        } else {
          autoType = "quant_simple";
        }
      }

      let defaultChart = "bar";
      if (autoType === "quant_grouped") defaultChart = "histogram";
      else if (autoType === "qualitative") defaultChart = "doughnut";

      return {
        index: colIdx,
        name: headerName || `Columna ${colIdx + 1}`,
        autoType,
        userType: autoType,
        chartType: defaultChart,
        validCount,
        emptyCount,
        invalidCount,
        uniqueCount: uniqueSet.size
      };
    });

    const firstActive = state.variables.find(v => v.userType !== "omit");
    state.selectedQuestionFilter = firstActive ? firstActive.index : "all";
    renderAll();
  }

  // ==========================================
  // 3. GENERACIÓN DE INTERPRETACIONES (10.1 y 10.2)
  // ==========================================

  function buildQuantitativeInterpretation(variable, stats) {
    const dispersionText = stats.cv < 15
      ? "muy homogénea (baja dispersión relativa)"
      : stats.cv < 30
        ? "moderadamente homogénea"
        : "heterogénea (alta variabilidad entre los encuestados)";

    const outlierText = stats.outliers.length === 0
      ? "No se detectaron valores atípicos (outliers) según el criterio de Tukey (1.5 × RIC)."
      : `Se identificaron ${stats.outliers.length} valor(es) atípico(s) fuera del intervalo [${stats.lowerFence}, ${stats.upperFence}]: (${stats.outliers.join(", ")}).`;

    return `
      En los <strong>${stats.n} registros válidos</strong> analizados para <em>"${variable.name}"</em>, el promedio (media aritmética) es de <strong>${stats.mean}</strong>, mientras que el 50% central de los encuestados se ubica alrededor de una mediana de <strong>${stats.median}</strong> y el valor más frecuente (moda) es <strong>${stats.modeText}</strong>.
      <br><br>
      Respecto a la dispersión, los datos oscilan entre un mínimo de <strong>${stats.min}</strong> y un máximo de <strong>${stats.max}</strong> (rango = <strong>${stats.range}</strong>), con una desviación estándar muestral de <strong>${stats.stdDev}</strong> (varianza = <strong>${stats.variance}</strong>) y un coeficiente de variación de <strong>${stats.cv}%</strong>, lo que indica una distribución <strong>${dispersionText}</strong>. El 50% intermedio de las respuestas se concentra entre <strong>Q1 = ${stats.q1}</strong> y <strong>Q3 = ${stats.q3}</strong> (RIC = <strong>${stats.iqr}</strong>), mientras que el 80% de los casos se sitúa entre el percentil 10 (<strong>${stats.p10}</strong>) y el percentil 90 (<strong>${stats.p90}</strong>).
      <br><br>
      En cuanto a la forma de la distribución, el coeficiente de asimetría (${stats.skewness}) muestra una distribución <strong>${stats.skewnessClass.toLowerCase()}</strong>, y la curtosis (${stats.kurtosisExcess}) refleja una forma <strong>${stats.kurtosisClass.toLowerCase()}</strong>. ${outlierText}
    `;
  }

  function buildQualitativeInterpretation(variable, stats) {
    const second = stats.freqRows[1];
    const secondPart = second
      ? `, seguida por <strong>"${second.label}"</strong> con <strong>${second.fi} respuestas (${second.pct}%)</strong>`
      : "";
    return `
      De un total de <strong>${stats.n} personas encuestadas</strong> en la pregunta cualitativa <em>"${variable.name}"</em>, la categoría predominante (moda cualitativa) es <strong>"${stats.topCategory.label}"</strong>, concentrando <strong>${stats.topCategory.fi} respuestas</strong> equivalentes al <strong>${stats.topCategory.pct}%</strong> del total${secondPart}. Por su parte, la opción con menor frecuencia fue <strong>"${stats.minCategory.label}"</strong> con <strong>${stats.minCategory.fi} menciones (${stats.minCategory.pct}%)</strong>. Esto evidencia una clara tendencia en la percepción de los estudiantes frente al uso de la Inteligencia Artificial en el ámbito educativo.
    `;
  }

  // ==========================================
  // 4. RENDERIZADO DE LA INTERFAZ
  // ==========================================

  function renderAll() {
    renderKPIs();
    renderValidationAndVariablesTable();
    renderQuestionsAnalysis();
    renderRawDataPreview();
    renderExecutiveReport();
  }

  function renderKPIs() {
    const activeVars = state.variables.filter(v => v.userType !== "omit");
    const quantVars = activeVars.filter(v => v.userType.startsWith("quant"));
    const qualVars = activeVars.filter(v => v.userType === "qualitative");
    const totalEmpties = state.variables.reduce((acc, v) => acc + v.emptyCount, 0);
    const totalInvalids = state.variables.reduce((acc, v) => acc + v.invalidCount, 0);

    document.getElementById("kpi-total-rows").textContent = state.cleanRows.length;
    document.getElementById("kpi-file-name").textContent = state.fileName;
    document.getElementById("kpi-total-vars").textContent = activeVars.length;
    document.getElementById("kpi-vars-breakdown").textContent =
      `${quantVars.length} Cuantitativas · ${qualVars.length} Cualitativas`;
    document.getElementById("kpi-data-quality").textContent =
      totalEmpties + totalInvalids === 0 ? "100% Válido" : `${totalEmpties} vacíos / ${totalInvalids} errores`;
    document.getElementById("kpi-duplicates").textContent =
      `${state.duplicateCount} duplicados exactos detectados`;
  }

  function renderValidationAndVariablesTable() {
    const tbody = document.getElementById("variables-table-body");
    tbody.innerHTML = "";

    state.variables.forEach((v, idx) => {
      const tr = document.createElement("tr");

      let badgeHtml = `<span class="badge badge-omit">Omitida (Marca temporal)</span>`;
      if (v.userType === "quant_simple") {
        badgeHtml = `<span class="badge badge-quant">Cuantitativa (Simple)</span>`;
      } else if (v.userType === "quant_grouped") {
        badgeHtml = `<span class="badge badge-grouped">Cuantitativa (Agrupada)</span>`;
      } else if (v.userType === "qualitative") {
        badgeHtml = `<span class="badge badge-qual">Cualitativa</span>`;
      }

      const statusHtml = (v.emptyCount === 0 && v.invalidCount === 0)
        ? `<span class="badge badge-ok">Sin errores (0 vacíos)</span>`
        : `<span class="badge badge-warn">${v.emptyCount} vacíos · ${v.invalidCount} no válidos</span>`;

      tr.innerHTML = `
        <td><strong>Col ${idx + 1}</strong></td>
        <td>${v.name}</td>
        <td>${badgeHtml}</td>
        <td>
          <select class="select-control var-type-select" data-index="${idx}">
            <option value="quant_simple" ${v.userType === "quant_simple" ? "selected" : ""}>Cuantitativa (Tabla Simple)</option>
            <option value="quant_grouped" ${v.userType === "quant_grouped" ? "selected" : ""}>Cuantitativa (Tabla Agrupada por Intervalos)</option>
            <option value="qualitative" ${v.userType === "qualitative" ? "selected" : ""}>Cualitativa (Categorías)</option>
            <option value="omit" ${v.userType === "omit" ? "selected" : ""}>Omitir del análisis (Ej. Marca temporal)</option>
          </select>
        </td>
        <td class="num-col">${v.validCount}</td>
        <td class="num-col">${v.uniqueCount}</td>
        <td>${statusHtml}</td>
      `;
      tbody.appendChild(tr);
    });

    document.querySelectorAll(".var-type-select").forEach(select => {
      select.addEventListener("change", e => {
        const idx = parseInt(e.target.getAttribute("data-index"), 10);
        const newType = e.target.value;
        state.variables[idx].userType = newType;
        if (newType === "quant_grouped") state.variables[idx].chartType = "histogram";
        else if (newType === "quant_simple") state.variables[idx].chartType = "bar";
        else if (newType === "qualitative") state.variables[idx].chartType = "doughnut";
        renderAll();
      });
    });
  }

  function destroyCharts() {
    Object.values(state.chartInstances).forEach(chart => {
      if (chart && typeof chart.destroy === "function") chart.destroy();
    });
    state.chartInstances = {};
  }

  function renderQuestionsAnalysis() {
    destroyCharts();
    const container = document.getElementById("questions-analysis-container");
    const filterBar = document.getElementById("question-filter-bar");
    container.innerHTML = "";
    filterBar.innerHTML = "";

    const activeVars = state.variables.filter(v => v.userType !== "omit");

    // Opción para ver todas las preguntas a la vez
    const allItem = document.createElement("button");
    allItem.className = `q-list-item ${state.selectedQuestionFilter === "all" ? "active" : ""}`;
    allItem.innerHTML = `
      <div class="q-item-num">ALL</div>
      <div class="q-item-body">
        <div class="q-item-title">Ver todas las preguntas</div>
        <div class="q-item-meta">${activeVars.length} variables en cascada</div>
      </div>
    `;
    allItem.addEventListener("click", () => {
      state.selectedQuestionFilter = "all";
      renderQuestionsAnalysis();
    });
    filterBar.appendChild(allItem);

    // Lista vertical de cada pregunta
    activeVars.forEach((v, qPos) => {
      const isQuant = v.userType.startsWith("quant");
      const item = document.createElement("button");
      item.className = `q-list-item ${state.selectedQuestionFilter === v.index ? "active" : ""}`;
      item.innerHTML = `
        <div class="q-item-num">P${qPos + 1}</div>
        <div class="q-item-body">
          <div class="q-item-title">${v.name.replace(/^\d+\.\s*/, "")}</div>
          <div class="q-item-meta">
            <span>${isQuant ? "Cuantitativa" : "Cualitativa"}</span>
            <span>·</span>
            <span>n = ${v.validCount}</span>
          </div>
        </div>
      `;
      item.addEventListener("click", () => {
        state.selectedQuestionFilter = v.index;
        renderQuestionsAnalysis();
      });
      filterBar.appendChild(item);
    });

    const varsToRender = state.selectedQuestionFilter === "all"
      ? activeVars
      : activeVars.filter(v => v.index === state.selectedQuestionFilter);

    varsToRender.forEach(variable => {
      const qPos = activeVars.findIndex(v => v.index === variable.index);
      const qNumber = qPos + 1;
      const prevVar = qPos > 0 ? activeVars[qPos - 1] : null;
      const nextVar = qPos < activeVars.length - 1 ? activeVars[qPos + 1] : null;

      if (variable.userType.startsWith("quant")) {
        renderQuantitativeCard(container, variable, qNumber, activeVars.length, prevVar, nextVar);
      } else {
        renderQualitativeCard(container, variable, qNumber, activeVars.length, prevVar, nextVar);
      }
    });
  }

  function buildStepNavHtml(qNumber, totalQuestions, prevVar, nextVar) {
    if (state.selectedQuestionFilter === "all") return "";
    return `
      <div class="q-step-nav no-print">
        <button class="btn prev-q-btn" ${!prevVar ? "disabled style='opacity:0.4;cursor:not-allowed;'" : ""}>
          ← Pregunta Anterior
        </button>
        <span style="font-size:0.82rem;font-weight:700;color:var(--text-muted);">
          Pregunta ${qNumber} de ${totalQuestions}
        </span>
        <button class="btn next-q-btn" ${!nextVar ? "disabled style='opacity:0.4;cursor:not-allowed;'" : ""}>
          Siguiente Pregunta →
        </button>
      </div>
    `;
  }

  function bindStepNavEvents(card, prevVar, nextVar) {
    const prevBtn = card.querySelector(".prev-q-btn");
    const nextBtn = card.querySelector(".next-q-btn");
    if (prevBtn && prevVar) {
      prevBtn.addEventListener("click", () => {
        state.selectedQuestionFilter = prevVar.index;
        renderQuestionsAnalysis();
      });
    }
    if (nextBtn && nextVar) {
      nextBtn.addEventListener("click", () => {
        state.selectedQuestionFilter = nextVar.index;
        renderQuestionsAnalysis();
      });
    }
  }

  function renderQuantitativeCard(container, variable, qNumber, totalQuestions, prevVar, nextVar) {
    const numValues = [];
    for (const row of state.cleanRows) {
      const val = parseNumericValue(row[variable.index], state.encodeYesNoAsNumeric);
      if (!Number.isNaN(val)) numValues.push(val);
    }

    const forceGrouped = variable.userType === "quant_grouped";
    const stats = analyzeQuantitative(numValues, forceGrouped);

    const card = document.createElement("div");
    card.className = "question-card";

    const freqTableHeaders = stats.useGrouped
      ? `
        <th>Intervalo [L<sub>i</sub>, L<sub>s</sub>)</th>
        <th class="num-col">Marca (x<sub>i</sub>)</th>
        <th class="num-col">f<sub>i</sub></th>
        <th class="num-col">F<sub>i</sub></th>
        <th class="num-col">h<sub>i</sub></th>
        <th class="num-col">H<sub>i</sub></th>
        <th class="num-col">%</th>
      `
      : `
        <th>Valor (x<sub>i</sub>)</th>
        <th class="num-col">Frec. Abs. (f<sub>i</sub>)</th>
        <th class="num-col">Frec. Acum. (F<sub>i</sub>)</th>
        <th class="num-col">Frec. Rel. (h<sub>i</sub>)</th>
        <th class="num-col">Rel. Acum. (H<sub>i</sub>)</th>
        <th class="num-col">Porcentaje (%)</th>
      `;

    const freqTableRows = stats.freqRows.map(r => stats.useGrouped ? `
      <tr>
        <td><strong>${r.label}</strong></td>
        <td class="num-col">${r.classMark}</td>
        <td class="num-col">${r.fi}</td>
        <td class="num-col">${r.Fi}</td>
        <td class="num-col">${r.hi.toFixed(4)}</td>
        <td class="num-col">${r.Hi.toFixed(4)}</td>
        <td class="num-col"><strong>${r.pct.toFixed(2)}%</strong></td>
      </tr>
    ` : `
      <tr>
        <td><strong>${r.label}</strong></td>
        <td class="num-col">${r.fi}</td>
        <td class="num-col">${r.Fi}</td>
        <td class="num-col">${r.hi.toFixed(4)}</td>
        <td class="num-col">${r.Hi.toFixed(4)}</td>
        <td class="num-col"><strong>${r.pct.toFixed(2)}%</strong></td>
      </tr>
    `).join("");

    const totalRowHtml = stats.useGrouped
      ? `<tr class="total-row"><td>Total</td><td class="num-col">—</td><td class="num-col">${stats.n}</td><td class="num-col">—</td><td class="num-col">1.0000</td><td class="num-col">—</td><td class="num-col">100.00%</td></tr>`
      : `<tr class="total-row"><td>Total</td><td class="num-col">${stats.n}</td><td class="num-col">—</td><td class="num-col">1.0000</td><td class="num-col">—</td><td class="num-col">100.00%</td></tr>`;

    card.innerHTML = `
      <div class="card-header">
        <div>
          <span class="badge ${stats.useGrouped ? "badge-grouped" : "badge-quant"}">
            Pregunta ${qNumber} · Variable Cuantitativa ${stats.useGrouped ? `(Agrupada en ${stats.kClasses} intervalos · Amplitud = ${stats.amplitude})` : "(Frecuencia Simple)"}
          </span>
          <h3 class="card-title" style="margin-top:8px;">${variable.name}</h3>
        </div>
        <div class="no-print" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
          <label style="font-size:0.8rem;font-weight:600;color:var(--text-muted);">Tabla:</label>
          <select class="select-control toggle-grouping-select" data-index="${variable.index}">
            <option value="quant_simple" ${!stats.useGrouped ? "selected" : ""}>Frecuencia Simple</option>
            <option value="quant_grouped" ${stats.useGrouped ? "selected" : ""}>Agrupada por Intervalos</option>
          </select>
          <label style="font-size:0.8rem;font-weight:600;color:var(--text-muted);">Gráfico:</label>
          <select class="select-control toggle-chart-select" data-index="${variable.index}">
            <option value="histogram" ${variable.chartType === "histogram" ? "selected" : ""}>Histograma + Polígono</option>
            <option value="bar" ${variable.chartType === "bar" ? "selected" : ""}>Diagrama de Barras</option>
            <option value="line" ${variable.chartType === "line" ? "selected" : ""}>Ojiva Acumulada</option>
            <option value="doughnut" ${variable.chartType === "doughnut" ? "selected" : ""}>Gráfico Circular (%)</option>
          </select>
        </div>
      </div>

      <!-- Cinta rápida de métricas destacadas -->
      <div class="quick-stats-ribbon">
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Media Aritmética (x̄)</div>
          <div class="quick-stat-value">${stats.mean}</div>
        </div>
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Mediana (Me)</div>
          <div class="quick-stat-value">${stats.median}</div>
        </div>
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Desv. Estándar (s)</div>
          <div class="quick-stat-value">${stats.stdDev}</div>
        </div>
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Coef. Variación (CV)</div>
          <div class="quick-stat-value">${stats.cv}%</div>
        </div>
      </div>

      <div class="question-analysis-grid">
        <div>
          <h4 style="font-size:0.82rem;margin-bottom:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.04em;">
            Distribución de Frecuencias (n = ${stats.n})
          </h4>
          <div class="table-responsive">
            <table>
              <thead><tr>${freqTableHeaders}</tr></thead>
              <tbody>${freqTableRows}${totalRowHtml}</tbody>
            </table>
          </div>
        </div>

        <div class="chart-box">
          <div class="chart-canvas-wrap">
            <canvas id="chart-var-${variable.index}"></canvas>
          </div>
        </div>
      </div>

      <div class="metrics-sections-grid">
        <div class="metric-group">
          <div class="metric-group-title">Tendencia Central</div>
          <div class="metric-row"><span class="metric-name">Media Aritmética (x̄)</span><span class="metric-val">${stats.mean}</span></div>
          <div class="metric-row"><span class="metric-name">Mediana (Me)</span><span class="metric-val">${stats.median}</span></div>
          <div class="metric-row"><span class="metric-name">Moda (Mo)</span><span class="metric-val">${stats.modeText}</span></div>
        </div>

        <div class="metric-group">
          <div class="metric-group-title">Variabilidad</div>
          <div class="metric-row"><span class="metric-name">Rango (R)</span><span class="metric-val">${stats.range}</span></div>
          <div class="metric-row"><span class="metric-name">Varianza Muestral (s²)</span><span class="metric-val">${stats.variance}</span></div>
          <div class="metric-row"><span class="metric-name">Desviación Estándar (s)</span><span class="metric-val">${stats.stdDev}</span></div>
          <div class="metric-row"><span class="metric-name">Coef. de Variación (CV)</span><span class="metric-val">${stats.cv}%</span></div>
          <div class="metric-row"><span class="metric-name">Rango Intercuartílico (RIC)</span><span class="metric-val">${stats.iqr}</span></div>
        </div>

        <div class="metric-group">
          <div class="metric-group-title">Posición (Cuartiles y Percentiles)</div>
          <div class="metric-row"><span class="metric-name">Mínimo / Máximo</span><span class="metric-val">${stats.min} / ${stats.max}</span></div>
          <div class="metric-row"><span class="metric-name">Cuartiles Q1 / Q2 / Q3</span><span class="metric-val">${stats.q1} | ${stats.q2} | ${stats.q3}</span></div>
          <div class="metric-row"><span class="metric-name">Percentil 10 (P10)</span><span class="metric-val">${stats.p10}</span></div>
          <div class="metric-row"><span class="metric-name">Percentiles P25 / P50 / P75</span><span class="metric-val">${stats.p25} | ${stats.p50} | ${stats.p75}</span></div>
          <div class="metric-row"><span class="metric-name">Percentil 90 (P90)</span><span class="metric-val">${stats.p90}</span></div>
        </div>

        <div class="metric-group">
          <div class="metric-group-title">Forma y Valores Atípicos</div>
          <div class="metric-row"><span class="metric-name">Coef. de Asimetría</span><span class="metric-val">${stats.skewness}</span></div>
          <div class="metric-row"><span class="metric-name">Simetría</span><span class="metric-val" style="color:var(--primary);text-align:right;">${stats.skewnessClass}</span></div>
          <div class="metric-row"><span class="metric-name">Curtosis (Exceso / Clásica)</span><span class="metric-val">${stats.kurtosisExcess} (${stats.kurtosisClassic})</span></div>
          <div class="metric-row"><span class="metric-name">Apuntamiento</span><span class="metric-val" style="color:var(--purple);text-align:right;">${stats.kurtosisClass}</span></div>
          <div class="metric-row"><span class="metric-name">Valores Atípicos</span><span class="metric-val">${stats.outliers.length === 0 ? "Ninguno (0)" : stats.outliers.join(", ")}</span></div>
        </div>
      </div>

      <div class="interpretation-box">
        <div class="interpretation-title">📊 Interpretación Estadística</div>
        <div class="interpretation-text">${buildQuantitativeInterpretation(variable, stats)}</div>
      </div>

      ${buildStepNavHtml(qNumber, totalQuestions, prevVar, nextVar)}
    `;

    container.appendChild(card);
    bindStepNavEvents(card, prevVar, nextVar);

    // Eventos de cambio de tabla / gráfico
    card.querySelector(".toggle-grouping-select").addEventListener("change", e => {
      variable.userType = e.target.value;
      renderAll();
    });
    card.querySelector(".toggle-chart-select").addEventListener("change", e => {
      variable.chartType = e.target.value;
      renderQuestionsAnalysis();
    });

    // Dibujar gráfico con Chart.js
    renderQuantitativeChart(variable, stats);
  }

  function renderQualitativeCard(container, variable, qNumber, totalQuestions, prevVar, nextVar) {
    const strValues = [];
    for (const row of state.cleanRows) {
      const cell = row[variable.index];
      if (cell !== null && cell !== undefined && String(cell).trim() !== "") {
        strValues.push(String(cell).trim());
      }
    }

    const stats = analyzeQualitative(strValues);
    const card = document.createElement("div");
    card.className = "question-card";

    const freqTableRows = stats.freqRows.map(r => `
      <tr>
        <td><strong>${r.label}</strong></td>
        <td class="num-col">${r.fi}</td>
        <td class="num-col">${r.Fi}</td>
        <td class="num-col">${r.hi.toFixed(4)}</td>
        <td class="num-col"><strong>${r.pct.toFixed(2)}%</strong></td>
      </tr>
    `).join("");

    card.innerHTML = `
      <div class="card-header">
        <div>
          <span class="badge badge-qual">Pregunta ${qNumber} · Variable Cualitativa</span>
          <h3 class="card-title" style="margin-top:8px;">${variable.name}</h3>
        </div>
        <div class="no-print" style="display:flex;gap:10px;align-items:center;">
          <label style="font-size:0.8rem;font-weight:600;color:var(--text-muted);">Gráfico:</label>
          <select class="select-control toggle-chart-select" data-index="${variable.index}">
            <option value="doughnut" ${variable.chartType === "doughnut" ? "selected" : ""}>Gráfico de Anillo / Pastel (%)</option>
            <option value="pie" ${variable.chartType === "pie" ? "selected" : ""}>Gráfico de Pastel Clásico</option>
            <option value="bar" ${variable.chartType === "bar" ? "selected" : ""}>Gráfico de Barras</option>
          </select>
        </div>
      </div>

      <div class="quick-stats-ribbon">
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Categoría Dominante (Moda)</div>
          <div class="quick-stat-value" style="font-size:1.05rem;">${stats.topCategory.label}</div>
        </div>
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Frecuencia Mayor (fᵢ)</div>
          <div class="quick-stat-value">${stats.topCategory.fi} (${stats.topCategory.pct}%)</div>
        </div>
        <div class="quick-stat-pill">
          <div class="quick-stat-label">Categorías Evaluadas</div>
          <div class="quick-stat-value">${stats.freqRows.length}</div>
        </div>
      </div>

      <div class="question-analysis-grid">
        <div>
          <h4 style="font-size:0.82rem;margin-bottom:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.04em;">
            Tabla de Frecuencias (n = ${stats.n})
          </h4>
          <div class="table-responsive">
            <table>
              <thead>
                <tr>
                  <th>Categoría / Respuesta</th>
                  <th class="num-col">Frec. Abs. (f<sub>i</sub>)</th>
                  <th class="num-col">Frec. Acum. (F<sub>i</sub>)</th>
                  <th class="num-col">Frec. Rel. (h<sub>i</sub>)</th>
                  <th class="num-col">Porcentaje (%)</th>
                </tr>
              </thead>
              <tbody>
                ${freqTableRows}
                <tr class="total-row">
                  <td>Total</td>
                  <td class="num-col">${stats.n}</td>
                  <td class="num-col">—</td>
                  <td class="num-col">1.0000</td>
                  <td class="num-col">100.00%</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <div class="chart-box">
          <div class="chart-canvas-wrap">
            <canvas id="chart-var-${variable.index}"></canvas>
          </div>
        </div>
      </div>

      <div class="interpretation-box">
        <div class="interpretation-title">📊 Interpretación Estadística</div>
        <div class="interpretation-text">${buildQualitativeInterpretation(variable, stats)}</div>
      </div>

      ${buildStepNavHtml(qNumber, totalQuestions, prevVar, nextVar)}
    `;

    container.appendChild(card);
    bindStepNavEvents(card, prevVar, nextVar);

    card.querySelector(".toggle-chart-select").addEventListener("change", e => {
      variable.chartType = e.target.value;
      renderQuestionsAnalysis();
    });

    renderQualitativeChart(variable, stats);
  }

  // ==========================================
  // 5. GRÁFICOS CON CHART.JS (Sección 11)
  // ==========================================

  function renderQuantitativeChart(variable, stats) {
    const canvas = document.getElementById(`chart-var-${variable.index}`);
    if (!canvas || typeof Chart === "undefined") return;

    const labels = stats.freqRows.map(r => r.label);
    const fiData = stats.freqRows.map(r => r.fi);
    const FiData = stats.freqRows.map(r => r.Fi);
    const chartType = variable.chartType || "bar";

    let config;
    if (chartType === "histogram") {
      config = {
        type: "bar",
        data: {
          labels,
          datasets: [
            {
              type: "line",
              label: "Polígono de Frecuencias (fi)",
              data: fiData,
              borderColor: "#f59e0b",
              backgroundColor: "#f59e0b",
              borderWidth: 2.5,
              tension: 0.25,
              pointRadius: 4
            },
            {
              type: "bar",
              label: "Frecuencia Absoluta (fi)",
              data: fiData,
              backgroundColor: "rgba(59, 76, 202, 0.72)",
              borderColor: "#3b4cca",
              borderWidth: 1.5,
              barPercentage: 1.0,
              categoryPercentage: 0.96
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            title: { display: true, text: "Histograma y Polígono de Frecuencias" },
            legend: { position: "bottom" }
          },
          scales: {
            y: { beginAtZero: true, title: { display: true, text: "Frecuencia Absoluta (fi)" } },
            x: { title: { display: true, text: stats.useGrouped ? "Intervalos de Clase" : "Valores (xi)" } }
          }
        }
      };
    } else if (chartType === "line") {
      config = {
        type: "line",
        data: {
          labels,
          datasets: [{
            label: "Frecuencia Acumulada (Fi) - Ojiva",
            data: FiData,
            borderColor: "#7c3aed",
            backgroundColor: "rgba(124, 58, 237, 0.16)",
            fill: true,
            tension: 0.25,
            pointRadius: 5
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            title: { display: true, text: "Ojiva de Frecuencia Acumulada (Fi)" },
            legend: { position: "bottom" }
          },
          scales: {
            y: { beginAtZero: true, title: { display: true, text: "Frecuencia Acumulada (Fi)" } }
          }
        }
      };
    } else if (chartType === "doughnut") {
      config = {
        type: "doughnut",
        data: {
          labels,
          datasets: [{
            data: fiData,
            backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length])
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            title: { display: true, text: "Distribución Porcentual" },
            legend: { position: "bottom" }
          }
        }
      };
    } else {
      config = {
        type: "bar",
        data: {
          labels,
          datasets: [{
            label: "Frecuencia Absoluta (fi)",
            data: fiData,
            backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length]),
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            title: { display: true, text: "Diagrama de Barras de Frecuencias" },
            legend: { display: false }
          },
          scales: {
            y: { beginAtZero: true, title: { display: true, text: "Encuestados (fi)" } },
            x: { title: { display: true, text: "Valor de la Variable" } }
          }
        }
      };
    }

    state.chartInstances[variable.index] = new Chart(canvas, config);
  }

  function renderQualitativeChart(variable, stats) {
    const canvas = document.getElementById(`chart-var-${variable.index}`);
    if (!canvas || typeof Chart === "undefined") return;

    const labels = stats.freqRows.map(r => `${r.label} (${r.pct}%)`);
    const data = stats.freqRows.map(r => r.fi);
    const chartType = variable.chartType || "doughnut";

    const config = {
      type: chartType === "bar" ? "bar" : chartType,
      data: {
        labels,
        datasets: [{
          label: "Frecuencia Absoluta (fi)",
          data,
          backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length]),
          borderRadius: chartType === "bar" ? 6 : 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          title: { display: true, text: "Distribución de Respuestas por Categoría" },
          legend: { position: "bottom", display: chartType !== "bar" }
        },
        scales: chartType === "bar" ? {
          y: { beginAtZero: true, title: { display: true, text: "Cantidad de Respuestas (fi)" } }
        } : {}
      }
    };

    state.chartInstances[variable.index] = new Chart(canvas, config);
  }

  // ==========================================
  // 6. VISTA DE DATOS CRUDOS Y REPORTE FINAL
  // ==========================================

  function renderRawDataPreview() {
    const thead = document.getElementById("raw-data-thead");
    const tbody = document.getElementById("raw-data-tbody");

    thead.innerHTML = `<tr><th>#</th>${state.headers.map(h => `<th>${h}</th>`).join("")}</tr>`;
    tbody.innerHTML = state.cleanRows.map((row, idx) => `
      <tr>
        <td><strong>${idx + 1}</strong></td>
        ${state.headers.map((_, cIdx) => `<td>${row[cIdx] !== undefined && row[cIdx] !== null ? row[cIdx] : ""}</td>`).join("")}
      </tr>
    `).join("");
  }

  function renderExecutiveReport() {
    const reportContainer = document.getElementById("executive-report-content");
    const activeVars = state.variables.filter(v => v.userType !== "omit");
    const quantVars = activeVars.filter(v => v.userType.startsWith("quant"));
    const qualVars = activeVars.filter(v => v.userType === "qualitative");

    const summaryRows = activeVars.map((v, idx) => {
      if (v.userType.startsWith("quant")) {
        const vals = state.cleanRows
          .map(r => parseNumericValue(r[v.index], state.encodeYesNoAsNumeric))
          .filter(x => !Number.isNaN(x));
        const st = analyzeQuantitative(vals, v.userType === "quant_grouped");
        return `
          <tr>
            <td><strong>P${idx + 1}</strong></td>
            <td>${v.name}</td>
            <td><span class="badge badge-quant">Cuantitativa</span></td>
            <td>Media = <strong>${st.mean}</strong> | Mediana = <strong>${st.median}</strong> | Moda = <strong>${st.modeText}</strong></td>
            <td>DE (s) = <strong>${st.stdDev}</strong> | CV = <strong>${st.cv}%</strong> | RIC = <strong>${st.iqr}</strong></td>
            <td>${st.skewnessClass} · ${st.kurtosisClass}</td>
          </tr>
        `;
      } else {
        const vals = state.cleanRows
          .map(r => r[v.index])
          .filter(x => x !== null && x !== undefined && String(x).trim() !== "");
        const st = analyzeQualitative(vals);
        return `
          <tr>
            <td><strong>P${idx + 1}</strong></td>
            <td>${v.name}</td>
            <td><span class="badge badge-qual">Cualitativa</span></td>
            <td colspan="2">Categoría principal: <strong>"${st.topCategory.label}"</strong> (${st.topCategory.fi} casos · <strong>${st.topCategory.pct}%</strong>)</td>
            <td>${st.freqRows.length} categorías evaluadas</td>
          </tr>
        `;
      }
    }).join("");

    reportContainer.innerHTML = `
      <div class="report-section">
        <h3>1. Tema y Problema Investigado</h3>
        <p><strong>Tema seleccionado:</strong> Uso de Inteligencia Artificial en la Educación.</p>
        <p><strong>Planteamiento del problema:</strong> El rápido crecimiento de las herramientas de Inteligencia Artificial (IA) ha transformado los hábitos de estudio en los estudiantes. Sin embargo, existe debate sobre el nivel real de familiaridad, el tiempo de uso semanal, la integración pedagógica en las instituciones y los dilemas éticos o de equidad académica que genera.</p>
      </div>

      <div class="report-section">
        <h3>2. Objetivos y Metodología Estadística</h3>
        <p><strong>Objetivo General:</strong> Aplicar los conceptos de Estadística Descriptiva mediante la recolección, organización, procesamiento, representación e interpretación de <strong>${state.cleanRows.length} encuestas válidas</strong> sobre el uso de la IA en la educación, utilizando este sistema automatizado.</p>
        <ul>
          <li><strong>Instrumento:</strong> Cuestionario estructurado de ${activeVars.length} preguntas (${quantVars.length} cuantitativas y ${qualVars.length} cualitativas).</li>
          <li><strong>Muestra procesada:</strong> n = ${state.cleanRows.length} estudiantes encuestados.</li>
          <li><strong>Depuración y validación:</strong> Se verificó la integridad de las columnas, registrando ${state.duplicateCount} duplicados y validando valores atípicos mediante el criterio intercuartílico de Tukey (1.5 × RIC).</li>
        </ul>
      </div>

      <div class="report-section">
        <h3>3. Cuadro Resumen de Hallazgos Estadísticos</h3>
        <div class="table-responsive">
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Pregunta / Variable</th>
                <th>Tipo</th>
                <th>Tendencia Central / Dominancia</th>
                <th>Variabilidad</th>
                <th>Forma / Distribución</th>
              </tr>
            </thead>
            <tbody>${summaryRows}</tbody>
          </table>
        </div>
      </div>

      <div class="report-section">
        <h3>4. Conclusiones Sustentadas en Evidencia Estadística</h3>
        <ul>
          <li><strong>Alta familiaridad y adopción:</strong> Los estudiantes muestran un nivel elevado de familiaridad con la IA (media superior a 4 en escala de 1 a 5) y utilizan en promedio entre 3 y 4 herramientas distintas en sus actividades académicas, dedicando alrededor de 7.5 horas semanales a su uso.</li>
          <li><strong>Impacto positivo frente a rezago institucional:</strong> Mientras que la valoración del impacto en el aprendizaje y progreso académico alcanza medias cercanas a 7.4 y 7.6 sobre 10, la percepción sobre qué tan bien integrada está la IA en el sistema educativo presenta una media menor (~3.0 sobre 5), evidenciando una brecha entre el uso estudiantil y la integración curricular formal.</li>
          <li><strong>Conciencia ética y cautela:</strong> La importancia de aplicar un enfoque ético y moral obtuvo la media más alta y la menor dispersión relativa (CV más bajo), mientras que en el ámbito cualitativo predomina el uso de la IA como <em>complemento</em> y ahorro de tiempo, acompañado de una postura de <em>preocupación</em> responsable sobre la equidad educativa.</li>
        </ul>
      </div>

      <div class="report-section">
        <h3>5. Recomendaciones</h3>
        <ul>
          <li><strong>Integración pedagógica guiada:</strong> Diseñar guías docentes claras que incorporen la IA como herramienta de complemento y pensamiento crítico en lugar de prohibirla sin lineamientos.</li>
          <li><strong>Talleres de uso ético y verificación:</strong> Dado el alto consenso sobre la importancia ética, implementar capacitaciones sobre citación de IA, validación de fuentes y prevención de dependencia tecnológica.</li>
          <li><strong>Evaluaciones adaptadas:</strong> Actualizar los métodos de evaluación académica para reducir la percepción de injusticia en el sistema educativo reportada en el estudio.</li>
        </ul>
      </div>
    `;
  }

  // ==========================================
  // 7. LECTURA DE ARCHIVOS (.XLSX Y .CSV)
  // ==========================================

  // Parser CSV RFC-4180 robusto (soporta comillas, comas internas y saltos)
  function parseCSVText(text) {
    const rows = [];
    let currentRow = [];
    let currentVal = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      const nextChar = text[i + 1];

      if (inQuotes) {
        if (char === '"' && nextChar === '"') {
          currentVal += '"';
          i++;
        } else if (char === '"') {
          inQuotes = false;
        } else {
          currentVal += char;
        }
      } else {
        if (char === '"') {
          inQuotes = true;
        } else if (char === ",") {
          currentRow.push(currentVal);
          currentVal = "";
        } else if (char === "\r" && nextChar === "\n") {
          currentRow.push(currentVal);
          rows.push(currentRow);
          currentRow = [];
          currentVal = "";
          i++;
        } else if (char === "\n" || char === "\r") {
          currentRow.push(currentVal);
          rows.push(currentRow);
          currentRow = [];
          currentVal = "";
        } else {
          currentVal += char;
        }
      }
    }
    if (currentVal !== "" || currentRow.length > 0) {
      currentRow.push(currentVal);
      rows.push(currentRow);
    }
    return rows;
  }

  function handleFileUpload(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = function (e) {
        const parsed = parseCSVText(e.target.result);
        if (parsed.length < 2) {
          alert("El archivo CSV no contiene suficientes filas.");
          return;
        }
        processDataset(parsed[0], parsed.slice(1), file.name);
      };
      reader.readAsText(file, "UTF-8");
    } else if (ext === "xlsx" || ext === "xls") {
      if (typeof XLSX === "undefined") {
        alert("La librería SheetJS (XLSX) no se pudo cargar. Por favor verifica tu conexión a internet o sube el archivo en formato .csv.");
        return;
      }
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });
        if (rows.length < 2) {
          alert("El archivo Excel no contiene suficientes registros.");
          return;
        }
        processDataset(rows[0], rows.slice(1), file.name);
      };
      reader.readAsArrayBuffer(file);
    } else {
      alert("Formato no soportado. Por favor selecciona un archivo .xlsx o .csv descargado de Google Forms.");
    }
  }

  // Exportar los datos actuales a CSV
  function downloadCurrentCSV() {
    const lines = [];
    lines.push(state.headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(","));
    for (const r of state.cleanRows) {
      lines.push(state.headers.map((_, idx) => {
        const v = r[idx] ?? "";
        if (typeof v === "number") return String(v);
        return `"${String(v).replace(/"/g, '""')}"`;
      }).join(","));
    }
    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "encuesta_ia_educacion_respuestas.csv";
    a.click();
    URL.revokeObjectURL(url);
  }

  // ==========================================
  // 8. INICIALIZACIÓN Y EVENTOS DOM
  // ==========================================

  document.addEventListener("DOMContentLoaded", () => {
    // Pestañas de navegación
    document.querySelectorAll(".tab-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
        btn.classList.add("active");
        const target = document.getElementById(btn.getAttribute("data-tab"));
        if (target) target.classList.add("active");
      });
    });

    // Subida de archivos (.xlsx y .csv)
    const fileInput = document.getElementById("file-input");
    const dropzone = document.getElementById("dropzone");

    dropzone.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", e => {
      if (e.target.files && e.target.files[0]) {
        handleFileUpload(e.target.files[0]);
      }
    });

    dropzone.addEventListener("dragover", e => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
    dropzone.addEventListener("dragleave", () => {
      dropzone.classList.remove("dragover");
    });
    dropzone.addEventListener("drop", e => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleFileUpload(e.dataTransfer.files[0]);
      }
    });

    // Botón recargar desde Google Sheets en vivo
    document.getElementById("btn-load-default").addEventListener("click", () => {
      syncFromGoogleSheet();
    });

    // Checkbox para codificar Sí/No en caso de preguntas binarias
    const checkEncode = document.getElementById("chk-encode-yesno");
    if (checkEncode) {
      checkEncode.addEventListener("change", e => {
        state.encodeYesNoAsNumeric = e.target.checked;
        processDataset(state.headers, state.rawRows, state.fileName);
      });
    }

    // Exportar o sincronizar en vivo desde Google Sheets (sin errores de CORS usando Google Visualization JSONP)
    function syncFromGoogleSheet() {
      const urlInput = document.getElementById("gsheet-url-input");
      const statusEl = document.getElementById("gsheet-sync-status");
      const rawUrl = urlInput
        ? urlInput.value.trim()
        : "https://docs.google.com/spreadsheets/d/1B2oywCb6BFszdzP0O0jtSbcAuoMOYqxOPKYMya78ebc/edit?usp=sharing";

      const idMatch = rawUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (!idMatch) {
        alert("Enlace de Google Sheets no válido. Asegúrate de pegar el link completo de tu hoja de cálculo.");
        return;
      }
      const sheetId = idMatch[1];
      const gidMatch = rawUrl.match(/[#&?]gid=([0-9]+)/);
      const gid = gidMatch ? gidMatch[1] : "0";

      if (statusEl) {
        statusEl.style.color = "var(--primary)";
        statusEl.textContent = "⏳ Conectando en tiempo real con Google Sheets...";
      }

      const callbackName = "__gsheetCallback_" + Date.now();
      const script = document.createElement("script");

      window[callbackName] = function (response) {
        delete window[callbackName];
        if (script.parentNode) script.parentNode.removeChild(script);

        if (!response || response.status !== "ok" || !response.table) {
          if (statusEl) {
            statusEl.style.color = "var(--danger)";
            statusEl.textContent = "❌ No se pudo leer la hoja. Verifica que tenga acceso público ('Cualquier persona con el enlace').";
          }
          return;
        }

        const cols = response.table.cols || [];
        const rawTableRows = response.table.rows || [];

        let headers = cols.map(c => (c && c.label ? String(c.label).trim() : ""));
        let dataRows = rawTableRows.map(r =>
          (r.c || []).map(cell => {
            if (!cell) return "";
            if (cell.f !== undefined && cell.f !== null) return cell.f;
            return cell.v !== undefined && cell.v !== null ? cell.v : "";
          })
        );

        // Si Google Visualization puso los encabezados en la primera fila de datos (cuando todas las columnas son texto)
        const hasAllEmptyLabels = headers.every(h => h === "");
        if (hasAllEmptyLabels && dataRows.length > 0) {
          headers = dataRows[0].map(h => String(h || "").trim());
          dataRows = dataRows.slice(1);
        }

        // Detectar automáticamente si las preguntas 1 a 8 todavía están como "Sí / No"
        let yesNoCount = 0;
        let totalChecked = 0;
        for (const r of dataRows) {
          for (let c = 1; c < headers.length; c++) {
            const val = String(r[c] || "").trim().toLowerCase();
            if (val !== "") {
              totalChecked++;
              if (val === "sí" || val === "si" || val === "no") yesNoCount++;
            }
          }
        }
        if (totalChecked > 0 && (yesNoCount / totalChecked) > 0.5) {
          state.encodeYesNoAsNumeric = true;
          if (checkEncode) checkEncode.checked = true;
        } else {
          state.encodeYesNoAsNumeric = false;
          if (checkEncode) checkEncode.checked = false;
        }

        processDataset(headers, dataRows, `Google Sheets en Vivo (${dataRows.length} respuestas reales)`);

        if (statusEl) {
          statusEl.style.color = "var(--success)";
          statusEl.innerHTML = `✅ <strong>¡Sincronizado en vivo!</strong> Se cargaron las <strong>${dataRows.length} respuestas reales</strong> actuales de tu Google Sheet.`;
        }
      };

      script.onerror = function () {
        delete window[callbackName];
        if (script.parentNode) script.parentNode.removeChild(script);
        if (statusEl) {
          statusEl.style.color = "var(--warning)";
          statusEl.textContent = `⚠️ Sin conexión a internet; mostrando las ${state.cleanRows.length} respuestas reales guardadas localmente.`;
        }
      };

      script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?gid=${gid}&tqx=responseHandler:${callbackName}`;
      document.body.appendChild(script);
    }

    const btnSyncTop = document.getElementById("btn-sync-gsheet-top");
    if (btnSyncTop) btnSyncTop.addEventListener("click", syncFromGoogleSheet);
    const btnSyncTab = document.getElementById("btn-sync-gsheet");
    if (btnSyncTab) btnSyncTab.addEventListener("click", syncFromGoogleSheet);

    // Alternar Modo Claro / Modo Oscuro
    const btnThemeToggle = document.getElementById("btn-theme-toggle");
    if (btnThemeToggle) {
      btnThemeToggle.addEventListener("click", () => {
        const html = document.documentElement;
        const isDark = html.getAttribute("data-theme") === "dark";
        const nextTheme = isDark ? "light" : "dark";
        html.setAttribute("data-theme", nextTheme);
        document.getElementById("theme-icon").textContent = nextTheme === "dark" ? "☀️" : "🌙";
        document.getElementById("theme-label").textContent = nextTheme === "dark" ? "Modo Claro" : "Modo Oscuro";
        if (typeof Chart !== "undefined") {
          Chart.defaults.color = nextTheme === "dark" ? "#cbd5e1" : "#475569";
          Chart.defaults.borderColor = nextTheme === "dark" ? "#1f2937" : "#e2e8f0";
        }
        renderQuestionsAnalysis();
      });
    }

    // Botones de descarga e impresión
    document.getElementById("btn-download-csv").addEventListener("click", downloadCurrentCSV);
    document.getElementById("btn-print-report").addEventListener("click", () => window.print());

    // 1. Cargar de inmediato las respuestas reales de respaldo
    if (window.ENCUESTA_PRECARGADA) {
      state.encodeYesNoAsNumeric = false;
      if (checkEncode) checkEncode.checked = false;
      processDataset(
        window.ENCUESTA_PRECARGADA.headers,
        window.ENCUESTA_PRECARGADA.rows,
        `Tabla Google Forms (${window.ENCUESTA_PRECARGADA.rows.length} respuestas reales)`
      );
    }

    // 2. Sincronizar automáticamente en vivo con Google Sheets al abrir la página
    syncFromGoogleSheet();
  });
})();
