using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToExcel
{
    class Property
    {
        private Excel.Range cellNom;
        private string nom;
        private Excel.Range cellValue;
        private string value;

        public string Nom { get => nom; set => nom = value; }
        public string Value { get => value; set => this.value = value; }
        public Excel.Range CellNom { get => cellNom; set => cellNom = value; }
        public Excel.Range CellValue { get => cellValue; set => cellValue = value; }

        public Property()
        { }

        public override string ToString()
        {
            return Nom + " " + CellNom.Address;
        }
    }
}
