using System;
using System.Collections.Generic;
using System.Text;

namespace ExportDataToExcel.Models
{
    public class ReportModel
    {
        public string WorkTime { get; set; }
        public string Date { get; set; }
        public string MasterName { get; set; }
        public List<Technique> Tecn { get; set; }
    }
    public class Technique
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string DriverName { get; set; }
        public string Poroda { get; set; }
        public string WorkPlace { get; set; }
        public bool IsWorking { get; set; }
        public List<Mashine> Mashines { get; set; }
    }

    public class Mashine
    {
        public int IdM { get; set; }
        public string Name { get; set; }
        public string DriverMName { get; set; }
        public string Reis { get; set; }
        public string Plecho { get; set; }
        public List<TechMin> TechMins { get; set; }

    }
    public class TechMin
    {
        public string Name { get; set; }
        public int Index { get; set; }
    }

}
