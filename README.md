# PruebaTecnica

## Query SQL para crear las tablas en POSTGRESQL
```sql
-- Tabla para almacenar las polilíneas
CREATE TABLE polylines (
    id serial PRIMARY KEY,
    geom geometry(LineString, 2236),
    layer text,
    srid text
);

-- Tabla para almacenar los bloques
CREATE TABLE blocks (
    id serial PRIMARY KEY,
    geom geometry(Point, 2236),
    layer text,
    srid text
);

```

