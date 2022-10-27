using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class ToCollectionSample
    {
        public static void Run()
        {
            using(var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Persons");
                // Load the sample data into the worksheet
                var range = ws.Cells["A1"].LoadFromCollection(ToCollectionSampleData.Persons, options =>
                {
                    options.PrintHeaders = true;
                    options.TableStyle = TableStyles.Dark1;
                });

                /**********************************************************
                 * ToCollection. Automaps cell data to class instance     *
                 **********************************************************/

                Console.WriteLine("******* Sample 33 - ToCollection ********\n");

                // export the data loaded into the worksheet above to a collection
                var exportedPersons = range.ToCollection<ToCollectionSamplePerson>();

                foreach(var person in exportedPersons)
                {
                    Console.WriteLine("***************************");
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}");
                    Console.WriteLine($"Height: {person.Height} cm");
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}");
                }

                Console.WriteLine();

                /**********************************************************
                 * ToCollectionWithMappings. Use this method to manually  *
                 * map all or just some of the cells to your class.       *
                 **********************************************************/

                Console.WriteLine("******* Sample 33 - ToCollectionWithMappings ********\n");

                var exportedPersons2 = ws.Cells["A1:D4"].ToCollectionWithMappings<ToCollectionSamplePerson>(
                    row => 
                    {
                        // this runs once per row in the range

                        // Create an instance of the exported class
                        var person = new ToCollectionSamplePerson();

                        // If some of the cells can be automapped, start by automapping the row data to the class
                        row.Automap(person);

                        // Note that you can only use column names as below
                        // if options.HeaderRow is set to the 0-based row index
                        // of the header row.
                        person.FirstName = row.GetValue<string>("FirstName");

                        // get value by the 0-based column index
                        person.Height = row.GetValue<int>(2);
                        
                        // return the class instance
                        return person;
                    }, 
                    options => options.HeaderRow = 0);

                foreach (var person in exportedPersons2)
                {
                    Console.WriteLine("***************************");
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}");
                    Console.WriteLine($"Height: {person.Height} cm");
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}");
                }

                Console.WriteLine();

                /**********************************************************
                 * ToCollection. Using property attributes for mappings,  *
                 * see the ToCollectionSamplePersonAttr class             *
                 **********************************************************/

                // Load the sample data into a new worksheet
                var ws2 = package.Workbook.Worksheets.Add("Ws2");
                var range2 = ws2.Cells["A1"].LoadFromCollection(ToCollectionSampleData.PersonsWithAttributes, options =>
                {
                    options.PrintHeaders = true;
                    options.TableStyle = TableStyles.Dark1;
                });

                Console.WriteLine("******* Sample 33 - ToCollection using attributes ********\n");

                // export the data loaded into the worksheet above to a collection
                var exportedPersons3 = range2.ToCollection<ToCollectionSamplePersonAttr>();

                foreach (var person in exportedPersons3)
                {
                    Console.WriteLine("***************************");
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}");
                    Console.WriteLine($"Height: {person.Height} cm");
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}");
                }

                Console.WriteLine();

                /**********************************************************
                 * ToCollection from a table                              *
                 **********************************************************/
                Console.WriteLine("******* Sample 33 - ToCollection from a table ********\n");
                // Load the sample data a new worksheet
                var ws3 = package.Workbook.Worksheets.Add("Ws3");
                var tableRange = ws3.Cells["A1"].LoadFromCollection(ToCollectionSampleData.Persons, options =>
                {
                    options.PrintHeaders = true;
                    options.TableStyle = TableStyles.Dark1;
                });
                var table = ws3.Tables.GetFromRange(tableRange);
                // export the data loaded into the worksheet above to a collection
                var exportedPersons4 = table.ToCollection<ToCollectionSamplePerson>();

                foreach (var person in exportedPersons4)
                {
                    Console.WriteLine("***************************");
                    Console.WriteLine($"Name: {person.FirstName} {person.LastName}");
                    Console.WriteLine($"Height: {person.Height} cm");
                    Console.WriteLine($"Birthdate: {person.BirthDate.ToShortDateString()}");
                }

                Console.WriteLine();
            }
        }
    }
}
