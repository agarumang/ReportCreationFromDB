using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace ReportGenerator.Services
{
    public class DatabaseService : IDatabaseService
    {
        private readonly string _connectionString;

        public DatabaseService()
        {
            var cs = ConfigurationManager.ConnectionStrings["DBconnection"];
            if (cs == null) throw new InvalidOperationException("Connection string 'DBconnection' not found in App.config.");
            _connectionString = cs.ConnectionString;
        }

        public async Task<List<string>> GetTablesAsync()
        {
            var list = new List<string>();
            using (var conn = new SqlConnection(_connectionString))
            {
                await conn.OpenAsync();
                var cmd = new SqlCommand(@"
                    SELECT TABLE_SCHEMA, TABLE_NAME
                    FROM INFORMATION_SCHEMA.TABLES
                    WHERE TABLE_TYPE = 'BASE TABLE'
                    ORDER BY TABLE_SCHEMA, TABLE_NAME", conn);

                using (var rdr = await cmd.ExecuteReaderAsync())
                {
                    while (await rdr.ReadAsync())
                    {
                        list.Add($"{rdr.GetString(0)}.{rdr.GetString(1)}");
                    }
                }
            }
            return list;
        }

        public async Task<List<ColumnDefinition>> GetColumnsAsync(string schema, string table)
        {
            var list = new List<ColumnDefinition>();
            using (var conn = new SqlConnection(_connectionString))
            {
                await conn.OpenAsync();
                var cmd = new SqlCommand(@"
                    SELECT COLUMN_NAME, DATA_TYPE
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_SCHEMA = @schema AND TABLE_NAME = @table
                    ORDER BY ORDINAL_POSITION", conn);
                cmd.Parameters.AddWithValue("@schema", schema);
                cmd.Parameters.AddWithValue("@table", table);

                using (var rdr = await cmd.ExecuteReaderAsync())
                {
                    while (await rdr.ReadAsync())
                    {
                        list.Add(new ColumnDefinition
                        {
                            Name = rdr.GetString(0),
                            DataType = rdr.GetString(1)
                        });
                    }
                }
            }
            return list;
        }

        public async Task<DataTable> GetColumnDataAsync(string schema, string table, string column, int top = 500)
        {
            var dt = new DataTable();
            string esc(string s) => s.Replace("]", "]]");
            var col = esc(column);
            var sch = esc(schema);
            var tbl = esc(table);

            var sql = $"SELECT TOP ({top}) [{col}] FROM [{sch}].[{tbl}]";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand(sql, conn))
            using (var adapter = new SqlDataAdapter(cmd))
            {
                await conn.OpenAsync();
                // adapter.Fill is synchronous and can block UI; run it on a background thread
                await Task.Run(() => adapter.Fill(dt));
            }
            return dt;
        }

        public async Task<DataTable> GetColumnsDataAsync(string schema, string table, IEnumerable<string> columns, int top = 500)
        {
            var cols = columns?.ToList() ?? new List<string>();
            if (cols.Count == 0) return new DataTable();

            string esc(string s) => s.Replace("]", "]]");
            var colList = string.Join(", ", cols.Select(c => $"[{esc(c)}]"));
            var sch = esc(schema);
            var tbl = esc(table);

            var sql = $"SELECT TOP ({top}) {colList} FROM [{sch}].[{tbl}]";

            var dt = new DataTable();
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand(sql, conn))
            using (var adapter = new SqlDataAdapter(cmd))
            {
                await conn.OpenAsync();
                // run the blocking Fill on a threadpool thread
                await Task.Run(() => adapter.Fill(dt));
            }
            return dt;
        }

        public async Task<DataTable> GetColumnsDataBetweenDatesAsync(string schema, string table, IEnumerable<string> columns, DateTime from, DateTime to)
        {
            var cols = columns?.ToList() ?? new List<string>();
            if (cols.Count == 0) return new DataTable();

            // Ensure sensible ordering of from/to
            if (from > to)
            {
                var tmp = from;
                from = to;
                to = tmp;
            }

            string esc(string s) => s.Replace("]", "]]");
            var colList = string.Join(", ", cols.Select(c => $"[{esc(c)}]"));
            var sch = esc(schema);
            var tbl = esc(table);

            // Use exclusive upper bound: [TimeStamp] >= @from AND [TimeStamp] < @to
            // Caller should pass to = (endDate.Date.AddDays(1)) to include the full end day.
            var sql = $"SELECT {colList} FROM [{sch}].[{tbl}] WHERE [TimeStamp] >= @from AND [TimeStamp] < @to ORDER BY [TimeStamp] ASC";

            var dt = new DataTable();
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand(sql, conn))
            using (var adapter = new SqlDataAdapter(cmd))
            {
                cmd.Parameters.Add(new SqlParameter("@from", SqlDbType.DateTime) { Value = from });
                cmd.Parameters.Add(new SqlParameter("@to", SqlDbType.DateTime) { Value = to });

                await conn.OpenAsync();
                // run the blocking Fill on a threadpool thread so UI can render the loader
                await Task.Run(() => adapter.Fill(dt));
            }
            return dt;
        }
    }
}