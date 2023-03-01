// See https://aka.ms/new-console-template for more information

using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using Dapper;

Console.WriteLine("Hello, World!");

await TestSqlServerAsync("system name", "server", "db", "", "");

await TestOracleAsync("system name", "", "", "", "", "");

await TestMySqlAsync("system name", "", "3306", "sys", "root", "", "SELECT @@version");

async Task TestSqlServerAsync(string systemFlag, string server, string db, string user, string password)
{
    try
    {
        using (var connection = new SqlConnection($"Server={server};Database={db};User Id={user};Password={password};"))
        {
            var result = await connection.QueryFirstAsync<string>("SELECT getdate()");
            Console.WriteLine($"{systemFlag} {server} ok " + result);
        }
    }
    catch (System.Exception ex)
    {
        Console.WriteLine($"{systemFlag} {server} conn error: " + ex.Message);
    }
}

Task TestOracleAsync(string systemFlag, string host, string port, string serviceName, string user, string password)
{
    try
    {
        string connectingStrings = @$"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)
            (HOST={host})(PORT={port})))(CONNECT_DATA=(SERVICE_NAME={serviceName})));User ID={user};Password={password};";

        using (OracleConnection connection = new OracleConnection(connectingStrings))
        {
            connection.Open();
            string sql = "select sysdate from dual";
            OracleCommand command = new OracleCommand(sql, connection);
            var obj = command.ExecuteScalar();
            Console.WriteLine($"{systemFlag} {host} OK " + obj);
        }
    }
    catch (System.Exception ex)
    {
        Console.WriteLine($"{systemFlag} {host} conn error: " + ex.Message);
    }

    return Task.CompletedTask;
}

async Task TestMySqlAsync(string systemFlag, string server, string port, string db, string user, string password, string? sql = null)
{
    try
    {
        using (var connection = new MySql.Data.MySqlClient.MySqlConnection($"server={server};username={user};pwd={password};port={port};database={db};SslMode=none;"))
        {
            var result = await connection.QueryFirstAsync<string>(sql ?? "SELECT now()");
            Console.WriteLine($"{systemFlag} {server} ok " + result);
        }
    }
    catch (System.Exception ex)
    {
        Console.WriteLine($"{systemFlag} {server} conn error: " + ex.Message);
    }
}


public class WzInfo
{
    public int id { get; set; }
    public string? code { get; set; }
    public string? mcgg { get; set; }
    public string? jldw { get; set; }

    public override string ToString()
    {
        return $"{id} {code} {mcgg} {jldw}";
    }
}

public class pkwz_lkpz_zw
{
    public string? gs { get; set; }
}