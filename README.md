# DotNETCoreAPIMapExcelToObject
A programming assignment for Opinity.
 
Assignment: (in dutch)
Een .NET Core (minimaal versie 2) API project met Swagger (UI) ondersteuning waarmee je (via Swagger) een Exel (XLSX) bestand kunt  uploaden.
Deze API moet het bestand vervolgens dynamisch kunnen omzetten naar een object structuur waarbij de mappings tussen Excel en object structuur simpel in code aan te passen moeten zijn (van kolom naar veld).
In deze API moet gebruikt gemaakt worden van Dependency Injection.
 
Project Source: https://github.com/martiendejong/DotNETCoreAPIMapExcelToObject 
 
Live Demo: http://mapexceltoobject.azurewebsites.net/ 
 
Namespace: ExcelToObject
 
Method 1: use an automapper profile
Add the mapping to the Automapper Profile:

	CreateMap<DataSet, DataTable>().ConstructUsing(dataSet => dataSet.Tables[""Blad1""]);
	CreateMap<DataRow, MappedObject>()
    .ForMember(mappedObject => mappedObject.Text, config => config.MapFrom(row => row[""Tekst""]))
    .ForMember(mappedObject => mappedObject.Number, config => config.MapFrom(row => row[""Getal""]))
    .ForMember(mappedObject => mappedObject.DecimalNumber, config => config.MapFrom(row => row[""Kommagetal""]))
    .ForMember(mappedObject => mappedObject.Date, config => config.MapFrom(row => row[""Datum""]))
    .ForMember(mappedObject => mappedObject.Id, config => config.MapFrom(row => row[""Guid""]));
 
Map an Excel file to a list of objects:

	var stream = File.OpenRead(""spreadsheat.xlsx"");
	var mapper = ExcelToObjectMapper<MappedObject>();
	IEnumerable<MappedObject> results = mapper.MapToObjects(stream);
 
 
Method 2: use a file named {ObjectType}.mapping

	AnotherMappedObject.mapping
	[Blad1]
	Text=Tekst
	Number=Getal
	DecimalNumber=KommaGetal
	Date=Datum
	Id=Guid
 
Map an Excel file to a list of objects:

	var stream = File.OpenRead(""spreadsheat.xlsx"");
	var mapper = ExcelToObjectConfigFileMapper<MappedObject>();
	IEnumerable<MappedObject> results = mapper.MapToObjects(stream);