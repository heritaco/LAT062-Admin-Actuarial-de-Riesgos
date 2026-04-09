# Proyecto Python - CreditMetrics para 3 bonos independientes

## Objetivo
Resolver el ejercicio de clase de **CreditMetrics** para un portafolio con tres bonos de calificacion inicial **A-**, **BBB-** y **B-**, todos con:

- vencimiento original de 5 anos,
- valor nominal de USD 100,
- cupones anuales de 7.5%, 9% y 11%, respectivamente,
- horizonte de riesgo de **1 ano**,
- probabilidades de migracion **independientes**.

El proyecto evita trabajar manualmente con las `18^3 = 5,832` combinaciones mediante una **convolucion discreta exacta** de las distribuciones marginales de cada bono, redondeadas a centavos.

---

## Supuestos implementados

### 1) Espacio de estados
Se usa el sistema de 18 estados mencionado en clase:

`AAA, AA+, AA, AA-, A+, A, A-, BBB+, BBB, BBB-, BB+, BB, BB-, B+, B, B-, CCC, D`

### 2) Matriz de transicion
Se usa la tabla de transicion de la diapositiva correspondiente a **Average One-Year Transition Rates For Global Corporates By Rating Modifier (1981-2021)**.

Como la tabla fuente trae una columna adicional `NR`, el proyecto ofrece dos politicas:

- `renormalize` (**default**): elimina `NR` y renormaliza las 18 probabilidades restantes para que sumen 1. Esta es la opcion coherente con el requerimiento de `18^3` estados.
- `keep_current`: asigna la masa de `NR` al rating actual.

### 3) Curva libre de riesgo
Se usa la curva del enunciado (04/08/2026):

- 1Y = 3.69%
- 2Y = 3.79%
- 3Y = 3.78%
- 5Y = 3.92%

Para 4Y se interpola con el promedio simple entre 3Y y 5Y, igual que en la plantilla Excel:

- 4Y = (3.78% + 3.92%) / 2 = **3.85%**

### 4) Curva por calidad crediticia
Se construye como:

\[
 y_t(r)=y_t^{rf}+s(r)
\]

donde `s(r)` es la sobretasa por rating. Para los notches intermedios se replica la logica de interpolacion de la plantilla Excel subida por el usuario.

### 5) Valuacion al horizonte
Para un bono que **no incumple** y migra al estado `r`, el valor a 1 ano es:

\[
V_1(r)=C+\sum_{u=1}^{3}\frac{C}{(1+y_u(r))^u}+\frac{C+N}{(1+y_4(r))^4}
\]

donde:

- `C` = cupon anual,
- `N` = valor nominal,
- quedan 4 anos residuales tras el horizonte de 1 ano.

Para `D` se usa el recovery indicado en el ejercicio:

- A-: 45%
- BBB-: 35%
- B-: 25%

### 6) Distribucion del portafolio
Si `X_i` es la distribucion discreta del valor del bono `i`, el valor del portafolio es:

\[
X_P=X_1+X_2+X_3
\]

Bajo independencia:

\[
\mathbb{P}(X_P=v)=\left(\mathbb{P}_{X_1} * \mathbb{P}_{X_2} * \mathbb{P}_{X_3}\right)(v)
\]

El codigo calcula esta convolucion en centavos y la **verifica** contra la enumeracion brute force de los 5,832 escenarios.

---

## Resultado principal (default: `--nr-policy renormalize`)

Al correr el proyecto, el resultado central es:

- **Valor actual del portafolio:** `327.223297`
- **Valor esperado a 1 ano:** `343.064302`
- **Perdida esperada:** `-15.841005`
  - signo negativo = ganancia esperada por carry/cupones
- **Volatilidad:** `23.425534`
- **VaR 99.9%:** `73.810000` aproximadamente a centavos
- **ES 99.9%:** alrededor de `93.59`
- **Cuantil inferior 0.1% del valor del portafolio:** `253.41`
- **Escenario umbral ilustrativo:** `A- | B+ | D`

Interpretacion breve:

> Con 99.9% de confianza, la perdida a 1 ano del portafolio no excede aproximadamente **USD 73.81** bajo el esquema CreditMetrics con migraciones independientes y con la politica `renormalize` para `NR`.

---

## Estructura

```text
creditmetrics_project/
|-- README.md
|-- pyproject.toml
|-- src/
|   `-- creditmetrics_project/
|       |-- __init__.py
|       |-- data.py
|       |-- model.py
|       `-- main.py
|-- tests/
|   `-- test_model.py
`-- outputs/
```

---

## Como correr

Primero, colocate en la carpeta raiz del proyecto:

```powershell
cd C:\Users\herie\Downloads\creditmetrics_project\mnt\data\creditmetrics_project
```

Si estas usando el layout `src/`, tienes dos formas correctas de ejecutar el proyecto.

### Opcion 0: instalar el proyecto una vez (recomendada)

```powershell
python -m pip install -e .
```

Luego ya puedes correr:

```powershell
python -m creditmetrics_project.main
```

o bien, usando el comando de consola:

```powershell
creditmetrics
```

Si PowerShell no reconoce `creditmetrics`, usa `python -m creditmetrics_project.main` o agrega tu carpeta `Scripts` de Python al `PATH`.

### Opcion 1: ejecutar directamente sin instalar

Desde la carpeta raiz del proyecto, en PowerShell:

```powershell
$env:PYTHONPATH = "src"
python -m creditmetrics_project.main
```

### Opcion 2: elegir politica para `NR`

```powershell
$env:PYTHONPATH = "src"
python -m creditmetrics_project.main --nr-policy keep_current
```

### Opcion 3: cambiar carpeta de salida

```powershell
$env:PYTHONPATH = "src"
python -m creditmetrics_project.main --output-dir outputs
```

### Error comun en Windows

Si ejecutas esto dentro de `src\creditmetrics_project`:

```powershell
python -m creditmetrics_project.main
```

fallara con `ModuleNotFoundError`, porque Python necesita que la carpeta visible sea la raiz del proyecto o que el paquete este instalado.

---

## Archivos generados

En `outputs/` se guardan:

- distribuciones por bono,
- distribucion del portafolio por convolucion,
- distribucion del portafolio por brute force,
- tabla comparativa entre convolucion y brute force,
- tabla completa de escenarios brute force,
- pasos intermedios de la convolucion,
- resumen JSON,
- metricas de cola,
- un archivo Excel `resultados_creditmetrics.xlsx` con hojas separadas,
- una carpeta `figuras/` con graficas y tablas renderizadas como imagenes,
- un PDF `reporte_creditmetrics.pdf` con las graficas y tablas acomodadas para entregar.

---

## Verificacion

El proyecto incluye una prueba que confirma que:

1. la convolucion discreta coincide con la enumeracion brute force a nivel centavos,
2. el brute force recorre exactamente `18^3 = 5,832` escenarios,
3. el VaR 99.9% es estable en aproximadamente **USD 73.81**.

Ejecutar pruebas:

```powershell
python -m pip install -e .[dev]
python -m pytest -q
```

---

## Comentario metodologico

Si: en este ejercicio **si conviene usar convolucion**. No elimina el hecho de que conceptualmente hay `18^3` estados, pero evita construir y ordenar manualmente una tabla gigante. Matematicamente, se reemplaza la enumeracion explicita por la suma de distribuciones discretas independientes.
