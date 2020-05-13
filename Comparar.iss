Sub Main
	Call CompareDatabase()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Archivo: Comparar bases de datos
Function CompareDatabase
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.CompareDB
	task.AddMatchKey "NUM_FACT", "NUM_FACT", "A"
	dbName = "Comparar_01.IMD"
	task.PerformTask dbName, "", "TOTAL", "TOTAL", "Ejemplox-Detalle de ventas.IMD"
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function