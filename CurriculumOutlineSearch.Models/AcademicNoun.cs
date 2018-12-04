using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CurriculumOutlineSearch.Models
{
    public class AcademicNounView
    {
        public int Id { get; set; }
        public string EnglishName { get; set; }
        public string ChineseName { get; set; }
        public int CategoryId { get; set; }
    }
}
