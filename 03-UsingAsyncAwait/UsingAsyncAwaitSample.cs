/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 *
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-SEP-2009
 *******************************************************************************/

using System;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Threading;
using OfficeOpenXml.Table;

namespace EPPlusSamples.SalesReport
{
    class UsingAsyncAwaitSample
    {
        /// <summary>
        /// Shows a few different ways to load / save asynchronous
        /// </summary>
        /// <param name="connectionString">The connection string to the SQLite database</param>
        public static async Task RunAsync(string connectionString)
        {
            var file = FileOutputUtil.GetFileInfo("03-AsyncAwait.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");

                using (var sqlConn = new SQLiteConnection(connectionString))
                {
                    sqlConn.Open();
                    using (var sqlCmd = new SQLiteCommand("select CompanyName, [Name], Email, Country, o.OrderId, orderdate, ordervalue, currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                    {
                        var range = await ws.Cells["B2"].LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), true, "Table1", TableStyles.Medium10);
                        range.AutoFitColumns();
                    }
                }

                await package.SaveAsync();
            }

            //Load the package async again.
            using (var package = new ExcelPackage())
            {
                await package.LoadAsync(file);

                var newWs = package.Workbook.Worksheets.Add("AddedSheet2");
                var range = await newWs.Cells["A1"].LoadFromTextAsync(FileInputUtil.GetFileInfo("03-UsingAsyncAwait", "Importfile.txt"), new ExcelTextFormat { Delimiter='\t' });
                range.AutoFitColumns();

                await package.SaveAsAsync(FileOutputUtil.GetFileInfo("03-AsyncAwait-LoadedAndModified.xlsx"));
            }
        }
    }
}
