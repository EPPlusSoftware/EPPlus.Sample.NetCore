using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class ToCollectionSamplePerson
    {
        public ToCollectionSamplePerson()
        {

        }

        public ToCollectionSamplePerson(string firstName, string lastName, int height, DateTime birthDate)
        {
            FirstName = firstName;
            LastName = lastName;
            Height = height;
            BirthDate = birthDate;
        }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public int Height { get; set; }

        public DateTime BirthDate { get; set; }
    }

    public class ToCollectionSamplePersonAttr
    {
        public ToCollectionSamplePersonAttr()
        {

        }

        public ToCollectionSamplePersonAttr(string firstName, string lastName, int height, DateTime birthDate)
        {
            FirstName = firstName;
            LastName = lastName;
            Height = height;
            BirthDate = birthDate;
        }

        [DisplayName("The persons first name")]
        public string FirstName { get; set; }

        [Description("The persons last name")]
        public string LastName { get; set; }

        [EpplusTableColumn(Header ="Height of the person")]
        public int Height { get; set; }

        public DateTime BirthDate { get; set; }
    }

    public static class ToCollectionSampleData
    {
        public static IEnumerable<ToCollectionSamplePerson> Persons
        {
            get
            {
                return new List<ToCollectionSamplePerson> 
                { 
                    new ToCollectionSamplePerson("John", "Doe", 176, new DateTime(1978, 3, 15)),
                    new ToCollectionSamplePerson("Sven", "Svensson", 183, new DateTime(1995, 11, 3)),
                    new ToCollectionSamplePerson("Jane", "Doe", 168, new DateTime(1989, 2, 26))
                };
            }
        }

        public static IEnumerable<ToCollectionSamplePersonAttr> PersonsWithAttributes
        {
            get
            {
                return new List<ToCollectionSamplePersonAttr>
                {
                    new ToCollectionSamplePersonAttr("John", "Doe", 176, new DateTime(1978, 3, 15)),
                    new ToCollectionSamplePersonAttr("Sven", "Svensson", 183, new DateTime(1995, 11, 3)),
                    new ToCollectionSamplePersonAttr("Jane", "Doe", 168, new DateTime(1989, 2, 26))
                };
            }
        }
    }
}
