using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToExcel
{
    class Local: Hierarchie
    {
        private int idRelationnel;
        private string nom;
        private int niveau;
        private Excel.Range cell;
        private List<Tuple<string, string>> properties = new List<Tuple<string, string>>();


        public int IdRelationnel { get => idRelationnel; set => idRelationnel = value; }
        public new string Nom { get => nom; set => nom = value; }
        public int Niveau { get => niveau; set => niveau = value; }
        public Excel.Range Cell { get => cell; set => cell = value; }
        public List<Tuple<string, string>> Properties { get => properties; set => properties = value; }

        public Local () { }

        public override string ToString ()
        {
            string fullPropertyString = null;
            foreach (Tuple<string, string> property in Properties)
            {
                fullPropertyString = fullPropertyString + "\n" + property.Item1 + ": " + property.Item2;
            }
            return "Type: Local\nID: " + idRelationnel + "\nNiveau: " + niveau + "\nNom: " + nom + "\n" + fullPropertyString;
        }


    }
}
