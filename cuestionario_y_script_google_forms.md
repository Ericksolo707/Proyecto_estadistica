# Cuestionario Rediseñado: Uso de Inteligencia Artificial en la Educación

Este cuestionario mantiene **exactamente las mismas 10 preguntas y temática** de tu encuesta original en Google Forms, pero ajustando las primeras 8 preguntas a **variables cuantitativas** (escalas de 1 a 5, escalas de 1 a 10 y valores numéricos de cantidad/horas) y conservando las últimas 2 como **variables cualitativas**, cumpliendo al 100% con la sección 6 de las instrucciones (**máximo 2 cualitativas y mínimo 8 cuantitativas**).

---

## Estructura de las 10 Preguntas

### Variables Cuantitativas (Preguntas 1 a 8)

1. **En una escala del 1 al 5, ¿qué tan familiarizado/a estás con el concepto de Inteligencia Artificial?**
   - *Tipo en Google Forms:* Escala lineal (1 = Nada familiarizado/a, 5 = Totalmente familiarizado/a)
   - *Tipo estadístico:* Cuantitativa discreta (Escala 1 a 5)
   - *Basada en tu pregunta original 1:* "¿Estás familiarizado/a con el concepto de Inteligencia Artificial?"

2. **¿Cuántas herramientas de Inteligencia Artificial conoces o tienes disponibles para tus estudios?**
   - *Tipo en Google Forms:* Respuesta corta (Número entero: 0, 1, 2, 3, 4, 5, 6...) o Selección (0 a 6+)
   - *Tipo estadístico:* Cuantitativa discreta (Conteo)
   - *Basada en tu pregunta original 2:* "¿Tienes acceso a herramientas de Inteligencia Artificial?"

3. **¿Cuántas horas a la semana utilizas herramientas con IA en tu proceso de aprendizaje?**
   - *Tipo en Google Forms:* Respuesta corta con validación numérica (0 a 25 horas)
   - *Tipo estadístico:* Cuantitativa continua / discreta con amplio recorrido (ideal para demostrar la **distribución de frecuencias agrupada por intervalos** que exige el punto 10.2.A)
   - *Basada en tu pregunta original 3:* "¿Has utilizado alguna herramienta que incorpore IA en tu proceso de aprendizaje?"

4. **En una escala del 1 al 10, ¿en qué medida consideras que la IA ha impactado tu habilidad de aprendizaje?**
   - *Tipo en Google Forms:* Escala lineal (1 = Impacto nulo/negativo, 10 = Impacto altamente positivo)
   - *Tipo estadístico:* Cuantitativa discreta (Escala 1 a 10)
   - *Basada en tu pregunta original 4:* "¿La IA ha afectado tu habilidad de aprendizaje?"

5. **En una escala del 1 al 5, ¿qué tan bien integrada consideras que está la IA en el sistema educativo?**
   - *Tipo en Google Forms:* Escala lineal (1 = Muy mal integrada, 5 = Excelentemente integrada)
   - *Tipo estadístico:* Cuantitativa discreta (Escala 1 a 5)
   - *Basada en tu pregunta original 5:* "¿Está bien integrada la IA en el sistema educativo?"

6. **En una escala del 1 al 5, ¿qué tan de acuerdo estás con que el uso de IA ha vuelto injusto el sistema educativo?**
   - *Tipo en Google Forms:* Escala lineal (1 = Totalmente en desacuerdo, 5 = Totalmente de acuerdo)
   - *Tipo estadístico:* Cuantitativa discreta (Escala 1 a 5)
   - *Basada en tu pregunta original 6:* "¿El uso de IA ha vuelto injusto el sistema educativo?"

7. **En una escala del 1 al 10, ¿qué tan esenciales son las herramientas de IA para tu progreso académico?**
   - *Tipo en Google Forms:* Escala lineal (1 = Nada esenciales, 10 = Totalmente indispensables)
   - *Tipo estadístico:* Cuantitativa discreta (Escala 1 a 10)
   - *Basada en tu pregunta original 7:* "¿Las herramientas de IA son esenciales para tu progreso académico?"

8. **En una escala del 1 al 5, ¿qué tan importante es que las herramientas de IA se usen con un enfoque ético y moral?**
   - *Tipo en Google Forms:* Escala lineal (1 = Nada importante, 5 = Sumamente importante)
   - *Tipo estadístico:* Cuantitativa discreta (Escala 1 a 5)
   - *Basada en tu pregunta original 8:* "¿Deberían las herramientas de IA usarse con un enfoque ético y moral?"

---

### Variables Cualitativas (Preguntas 9 y 10 — Conservadas idénticas a tu encuesta)

9. **¿Por qué recurres a herramientas de IA?**
   - *Opciones:* `Facilita procesos` | `Ahorra tiempo` | `Sirve de complemento`
   - *Tipo estadístico:* Cualitativa nominal

10. **¿Cómo te sientes al usar IA en tus actividades escolares?**
    - *Opciones:* `Despreocupado/a` | `Preocupado/a` | `Fenomenal`
    - *Tipo estadístico:* Cualitativa nominal / ordinal

---

## (Opcional) Cómo crear el Google Form y la Hoja de Cálculo en tu cuenta en 1 minuto

Ya tienes generados en esta misma carpeta los archivos listos para usar:
- `encuesta_ia_educacion_respuestas.xlsx` (con 75 respuestas completas)
- `encuesta_ia_educacion_respuestas.csv` (con 75 respuestas completas)

Si además quieres tener el **Formulario de Google Forms nuevo** y su **Google Sheet vinculada** en tu cuenta de Google sin hacerlo a mano pregunta por pregunta:

1. Entra a [https://script.google.com](https://script.google.com) y haz clic en **"Nuevo proyecto"**.
2. Borra el código que aparece y pega el siguiente bloque:

```javascript
function crearFormularioYTablaEstadistica() {
  var form = FormApp.create('Uso De Inteligencia Artificial En La Educación');
  form.setDescription('Este cuestionario tiene como propósito conocer la percepción, el acceso y el impacto del uso de las herramientas de Inteligencia Artificial en el ámbito educativo actual.');

  form.addScaleItem()
    .setTitle('1. En una escala del 1 al 5, ¿qué tan familiarizado/a estás con el concepto de Inteligencia Artificial?')
    .setBounds(1, 5)
    .setLabels('Nada familiarizado/a', 'Totalmente familiarizado/a')
    .setRequired(true);

  form.addScaleItem()
    .setTitle('2. ¿Cuántas herramientas de Inteligencia Artificial conoces o tienes disponibles para tus estudios?')
    .setBounds(1, 5)
    .setLabels('1 herramienta', '5 o más herramientas')
    .setRequired(true);

  var q3 = form.addTextItem()
    .setTitle('3. ¿Cuántas horas a la semana utilizas herramientas con IA en tu proceso de aprendizaje?')
    .setRequired(true);
  var numValidation = FormApp.createTextValidation().requireNumberBetween(0, 40).build();
  q3.setValidation(numValidation);

  form.addScaleItem()
    .setTitle('4. En una escala del 1 al 10, ¿en qué medida consideras que la IA ha impactado tu habilidad de aprendizaje?')
    .setBounds(1, 10)
    .setLabels('Impacto muy bajo', 'Impacto muy alto')
    .setRequired(true);

  form.addScaleItem()
    .setTitle('5. En una escala del 1 al 5, ¿qué tan bien integrada consideras que está la IA en el sistema educativo?')
    .setBounds(1, 5)
    .setLabels('Muy mal integrada', 'Muy bien integrada')
    .setRequired(true);

  form.addScaleItem()
    .setTitle('6. En una escala del 1 al 5, ¿qué tan de acuerdo estás con que el uso de IA ha vuelto injusto el sistema educativo?')
    .setBounds(1, 5)
    .setLabels('Totalmente en desacuerdo', 'Totalmente de acuerdo')
    .setRequired(true);

  form.addScaleItem()
    .setTitle('7. En una escala del 1 al 10, ¿qué tan esenciales son las herramientas de IA para tu progreso académico?')
    .setBounds(1, 10)
    .setLabels('Nada esenciales', 'Totalmente esenciales')
    .setRequired(true);

  form.addScaleItem()
    .setTitle('8. En una escala del 1 al 5, ¿qué tan importante es que las herramientas de IA se usen con un enfoque ético y moral?')
    .setBounds(1, 5)
    .setLabels('Nada importante', 'Sumamente importante')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('9. ¿Por qué recurres a herramientas de IA?')
    .setChoiceValues(['Facilita procesos', 'Ahorra tiempo', 'Sirve de complemento'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('10. ¿Cómo te sientes al usar IA en tus actividades escolares?')
    .setChoiceValues(['Despreocupado/a', 'Preocupado/a', 'Fenomenal'])
    .setRequired(true);

  Logger.log('Formulario creado: ' + form.getPublishedUrl());
  Logger.log('Editar formulario: ' + form.getEditUrl());
}
```

3. Haz clic en **Ejecutar** (`Run`) y en los registros te dará el enlace de tu nuevo formulario listo. Para la hoja de cálculo, también puedes subir directamente el archivo `encuesta_ia_educacion_respuestas.xlsx` a tu Google Drive.
