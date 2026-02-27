using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace ReportGenerator.Services
{
    public interface IDatabaseService
    {
        Task<List<string>> GetTablesAsync();
        Task<List<ColumnDefinition>> GetColumnsAsync(string schema, string table);
        Task<DataTable> GetColumnDataAsync(string schema, string table, string column, int top = 500);
        Task<DataTable> GetColumnsDataAsync(string schema, string table, IEnumerable<string> columns, int top = 500);
        Task<DataTable> GetColumnsDataBetweenDatesAsync(string schema, string table, IEnumerable<string> columns, System.DateTime from, System.DateTime to);
    }
}
