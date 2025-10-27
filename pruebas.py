import pyodbc

Ruta = r".\Almacen1.accdb"
Conexion = (r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            rf"DBQ={Ruta};"
)

try:
    conn=pyodbc.connect(Conexion)
    cursor=conn.cursor()

    for fila in cursor.tables(tableType='TABLE'):
        print(fila.table_name)

    cursor.execute("SELECT Productos.NombreProducto, Proveedores.NombreCompañía, Proveedores.Teléfono, Categorías.NombreCategoría, Productos.UnidadesEnExistencia, Productos.UnidadesEnPedido FROM (Productos LEFT JOIN Proveedores ON Productos.IdProveedor = Proveedores.IdProveedor) LEFT JOIN Categorías ON Productos.IdCategoría = Categorías.IdCategoría WHERE UnidadesEnExistencia > 100;")
    
    columnas=[col[0] for col in cursor.description]
    print("\t".join(columnas))
        
    for r in cursor.fetchall():
        print("\t".join(str(v) if v is not None else "" for v in r))

finally:
    cursor.close()
    conn.close()
    