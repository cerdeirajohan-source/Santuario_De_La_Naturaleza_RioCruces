# Geovisualizador · Santuario de la Naturaleza Río Cruces

App en Streamlit + Folium con los proyectos SEIA del área de influencia del santuario (Valdivia y Mariquina).

## Estructura del repositorio
```
app.py
requirements.txt
data/
  Fuentes_y_proyectos_Santuario_Rio_Cruces.xlsx
  area_influencia_conaf2019_aprox.geojson
  santuario_rio_cruces.geojson   (opcional: límite real; si falta se usa uno provisorio)
```

## Ejecutar en local
```
pip install -r requirements.txt
streamlit run app.py
```

## Publicar
1. Sube todo a GitHub (repo `Santuario_RioCruces`).
2. En https://share.streamlit.io → *New app* → elige el repo, rama `main`, archivo `app.py`.
