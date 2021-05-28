using System;
using System.Windows;
using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using PARTITF;


namespace MinimalCatia
{
    class CatiaConnection
    {
        INFITF.Application hsp_catiaApp;
        MECMOD.PartDocument hsp_catiaPartDoc;
        MECMOD.Sketch hsp_catiaSkizze;

        ShapeFactory SF;
        HybridShapeFactory HSF;

        Part myPart;
        Sketches mySketches;

        public bool CATIALaeuft()
        {
            try
            {
                object catiaObject = System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "CATIA.Application");
                hsp_catiaApp = (INFITF.Application) catiaObject;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public Boolean ErzeugePart()
        {
            INFITF.Documents catDocuments1 = hsp_catiaApp.Documents;
            hsp_catiaPartDoc = catDocuments1.Add("Part") as MECMOD.PartDocument;
            myPart = hsp_catiaPartDoc.Part;

            return true;
        }

        public void ErstelleLeereSkizze()
        {
            // Factories für das Erzeugen von Modellelementen (Std und Hybrid)
            SF = (ShapeFactory)myPart.ShapeFactory;
            HSF = (HybridShapeFactory)myPart.HybridShapeFactory;

            // geometrisches Set auswaehlen und umbenennen
            HybridBodies catHybridBodies1 = myPart.HybridBodies;
            HybridBody catHybridBody1;
            try
            {
                catHybridBody1 = catHybridBodies1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                MessageBox.Show("Kein geometrisches Set gefunden! " + Environment.NewLine +
                    "Ein PART manuell erzeugen und ein darauf achten, dass 'Geometisches Set' aktiviert ist.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            catHybridBody1.set_Name("Profile");

            // neue Skizze im ausgewaehlten geometrischen Set auf eine Offset Ebene legen
            mySketches = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = myPart.OriginElements;
            HybridShapePlaneOffset hybridShapePlaneOffset1 = HSF.AddNewPlaneOffset(
                (Reference)catOriginElements.PlaneYZ, 90.000000, false);
            hybridShapePlaneOffset1.set_Name("OffsetEbene");
            catHybridBody1.AppendHybridShape(hybridShapePlaneOffset1);
            myPart.InWorkObject = hybridShapePlaneOffset1;
            myPart.Update();

            HybridShapes hybridShapes1 = catHybridBody1.HybridShapes;
            Reference catReference1 = (Reference)hybridShapes1.Item("OffsetEbene");

            hsp_catiaSkizze = mySketches.Add(catReference1);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystem();

            // Part aktualisieren
            myPart.Update();
        }

        private void ErzeugeAchsensystem()
        {
            object[] arr = new object[] {0.0, 0.0, 0.0,
                                         0.0, 1.0, 0.0,
                                         0.0, 0.0, 1.0 };
            hsp_catiaSkizze.SetAbsoluteAxisData(arr);
        }

        public void ErzeugeProfil(Double b, Double h)
        {
            // Skizze umbenennen
            hsp_catiaSkizze.set_Name("Rechteck");

            // Rechteck in Skizze einzeichnen
            // Skizze oeffnen
            Factory2D catFactory2D1 = hsp_catiaSkizze.OpenEdition();

            // Rechteck erzeugen

            // erst die Punkte
            Point2D catPoint2D1 = catFactory2D1.CreatePoint(-50, 50);
            Point2D catPoint2D2 = catFactory2D1.CreatePoint(50, 50);
            Point2D catPoint2D3 = catFactory2D1.CreatePoint(50, -50);
            Point2D catPoint2D4 = catFactory2D1.CreatePoint(-50, -50);

            // dann die Linien
            Line2D catLine2D1 = catFactory2D1.CreateLine(-50, 50, 50, 50);
            catLine2D1.StartPoint = catPoint2D1;
            catLine2D1.EndPoint = catPoint2D2;

            Line2D catLine2D2 = catFactory2D1.CreateLine(50, 50, 50, -50);
            catLine2D2.StartPoint = catPoint2D2;
            catLine2D2.EndPoint = catPoint2D3;

            Line2D catLine2D3 = catFactory2D1.CreateLine(50, -50, -50, -50);
            catLine2D3.StartPoint = catPoint2D3;
            catLine2D3.EndPoint = catPoint2D4;

            Line2D catLine2D4 = catFactory2D1.CreateLine(-50, -50, -50, 50);
            catLine2D4.StartPoint = catPoint2D4;
            catLine2D4.EndPoint = catPoint2D1;

            // Skizzierer verlassen
            hsp_catiaSkizze.CloseEdition();
            // Part aktualisieren
            myPart.Update();
        }

        public void ErzeugeBalken(Double l)
        {
            // Hauptkoerper in Bearbeitung definieren
            myPart.InWorkObject = myPart.MainBody;

            // Block(Balken) erzeugen
            Pad catPad1 = SF.AddNewPad(hsp_catiaSkizze, l);

            // Block umbenennen
            catPad1.set_Name("Balken");

            // Part aktualisieren
            myPart.Update();
        }



    }
}
