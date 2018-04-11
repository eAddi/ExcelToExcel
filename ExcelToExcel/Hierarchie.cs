using System;
using System.Collections.Generic;
using System.Threading;

namespace ExcelToExcel
{
    class Hierarchie
    {
        private static int _globalCount;
        private int idEntite;

        private int id;
        private string nom;
        private List<Hierarchie> list = new List<Hierarchie>();

        private int niveauHierarchie;
        private string niveauH;

        public int Id { get => id; set => id = value; }
        public string Nom { get => nom; set => nom = value; }
        public List<Hierarchie> List { get => list; set => list = value; }
        public int NiveauHierarchie { get => niveauHierarchie; set => niveauHierarchie = value; }
        public string NiveauH { get => niveauH; set => niveauH = value; }
        public int IdEntite { get => idEntite; set => idEntite = value; }

        public Hierarchie() { }

        public Hierarchie(string nom, int id, int niveauHierarchie, List<Hierarchie> list)
        {
            this.nom = nom;
            this.niveauHierarchie = niveauHierarchie;
            this.list = list;
            
            if (niveauHierarchie == 0)
            {
                niveauH = "Entité";
                IdEntite = Interlocked.Increment(ref _globalCount);
                
            }
            else
            {
                this.id = id;
                switch (niveauHierarchie)
                {
                    case 1:
                        niveauH = "Secteur";
                        break;
                    case 2:
                        niveauH = "Sous-Secteur";
                        break;
                    case 3:
                        niveauH = "Zone";
                        break;
                    case 4:
                        niveauH = "Sous-Zone";
                        break;
                    case 5:
                        niveauH = "Groupe";
                        break;
                    case 6:
                        niveauH = "Sous-Groupe";
                        break;
                    default:
                        niveauH = "Sous-Groupe/.../Local";
                        break;
                }
            }


        }


        public override string ToString()
        {
            if(niveauH == "Entité")
            {
                return "Type: " + niveauH + "\nID: " + idEntite + "\nNiveau: " + niveauHierarchie + "\nNom: " + nom;
            }
            else
            {
                return "Type: " + niveauH + "\nID: " + id + "\nNiveau: " + niveauHierarchie + "\nNom: " + nom;
            }  
        }

        public string DumpH()
        {
            if (this is Local l)
            {
                return l.ToString() + "\n";
            }
            else
            {
                String toPrint = ToString() + "{\n";
                foreach (Hierarchie ssH in list)
                {
                    toPrint += ssH.DumpH();
                }
                toPrint += "}\n";
                return toPrint;
            }
        }


    }
}
