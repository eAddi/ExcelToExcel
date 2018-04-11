using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToExcel
{
    class Cell_Style: IEquatable<Cell_Style>
    {
        private Color cellColor;
        private Color fontColor;
        private int fontSize;
        private bool underline;
        private bool bold;
        private bool italic;


        public Cell_Style(Excel.Range cell)
        {
            cellColor = ColorTranslator.FromOle(Convert.ToInt32(cell.Interior.Color));
            underline = (cell.Font.Underline == (int)Excel.XlUnderlineStyle.xlUnderlineStyleSingle) ? true : false;
            bold = cell.Font.Bold;
            fontColor = ColorTranslator.FromOle(Convert.ToInt32(cell.Font.Color));
            fontSize = Convert.ToInt32(cell.Font.Size);
            italic = cell.Font.Italic;
        }

        public int FontSize { get => fontSize; set => fontSize = value; }
        public Color FontColor { get => fontColor; set => fontColor = value; }
        public Color CellColor { get => cellColor; set => cellColor = value; }
        public bool Underline { get => underline; set => underline = value; }
        public bool Bold { get => bold; set => bold = value; }
        public bool Italic { get => italic; set => italic = value; }

        public override string ToString()
        {
            return "Cell color: " + cellColor.ToString() + "\n" +
                "Font color: " + fontColor.ToString() + "\n" +
                "Font size: " + fontSize.ToString() + "\n" +
                "Underline: " + underline.ToString() + "\n" +
                "Bold: " + bold.ToString() + "\n" +
                "Italic : " + italic.ToString() + "\n";
        }

        public bool Equals(Cell_Style other) 
        {
            if (other.Bold == bold &&
                other.CellColor == cellColor &&
                other.FontColor == fontColor &&
                other.FontSize == fontSize &&
                other.Underline == underline && 
                other.Italic == italic)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
