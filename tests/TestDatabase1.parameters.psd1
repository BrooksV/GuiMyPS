@{
    "ConnectionString" = "Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename=TestDatabase1.mdf;Integrated Security=True";
    "TableName" = "DBTable";
    "SqlServer" = "(localdb)\MSSQLLocalDB";
    "Database" = "TestDatabase1.mdf";
    "Query" = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = `'BASE TABLE`' ORDER BY TABLE_SCHEMA, TABLE_NAME;"
}
