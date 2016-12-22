using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

using App = Autodesk.AutoCAD.ApplicationServices;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Gem = Autodesk.AutoCAD.Geometry;
using Rtm = Autodesk.AutoCAD.Runtime;
using Win = Autodesk.Windows;

[assembly: Rtm.CommandClass(typeof(Opening_testLevel.Commands))]


namespace Opening_testLevel
{

    //internal class ErroeMetric
    //{
    //    public string ErrorName;
    //    public int ErrorCount;
    //    public short ErrorColorIndex;
    //}


    public class Commands : Rtm.IExtensionApplication
    {

        public void Initialize()
        {
            //Ed.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //ed.WriteMessage("\ninitialization test start...");

            // даем команду отслеживать изменения рабочей области (Workspace) AutoCAD 
            Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanged +=
                new Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventHandler(onSystemVariableChanged);

            //String TabName = "ACAD_DLL_Ribbon";
            //String TabTitle = "ACAD_DLL";
            //String PanelName = "opening";
            //String buttonName = "testLevel";
            //String _command = "._bx_opening_testLevel ";
            //////acDoc.SendStringToExecute("._circle 2,2,0 4 ", true, false, false);
            //AddRibbons.AddRibbon(TabName, TabTitle, PanelName, buttonName, _command);

        }
        public void Terminate()
        {
            Console.WriteLine("finish!");
        }


        // обработчик изменения рабочей области (Workspace) AutoCAD
        void onSystemVariableChanged(object sender, Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventArgs e)
        {
            // если рабочая область изменилась и в новой области есть лента (Ribbon)
            if ((e.Name == "WSCURRENT") && (Win.ComponentManager.Ribbon != null))
            {
                // создаем вкладку
                //addMyRibbonTab();
                // AddRibbons.AddRibbon(TabName, TabTitle, PanelName, buttonName, _command);
            }
        }



        //Проверка блоков отверстий, в Autocad, на попадание в отметку этажа
        [Rtm.CommandMethod("bx_opening_testLevel")]
        static public void bx_opening_testLevel()
        {

            //string ver = My.Application.Info.Version.ToString;
            string ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            //string ver = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString();
            if (sec.CheckVER("http://experement.spb.ru/", ver, "bx_opening_testLevel") == 1 & sec.coat == 21)
            {
                //Дополнительная проверка, может быть раскидана по коду.
                //If coat <> 21 Then Exit Sub
            }
            else
            {
                return;
            }


            // Получение текущего документа и базы данных
            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            Db.Database acCurDb = acDoc.Database;
            Ed.Editor acEd = acDoc.Editor;


            //Ошибки
            Dictionary<string, ErroeMetric> outError = new Dictionary<string, ErroeMetric>();
            //Список отверстий
            List<Opening> openingList = new List<Opening>();


            Ed.PromptDoubleOptions DownLevelOpt = new Ed.PromptDoubleOptions("\n введи нижнюю отметку стены: ");
            DownLevelOpt.AllowNone = false;
            Ed.PromptDoubleResult DownLevelRes = acEd.GetDouble(DownLevelOpt);
            if (DownLevelRes.Status != Ed.PromptStatus.OK)
                return;


            Ed.PromptDoubleOptions UpLevelOpt = new Ed.PromptDoubleOptions("\n введи верхнюю отметку стены: ");
            UpLevelOpt.AllowNone = false;
            Ed.PromptDoubleResult UpLevelRes = acEd.GetDouble(UpLevelOpt);
            if (UpLevelRes.Status != Ed.PromptStatus.OK)
                return;


            if (DownLevelRes.Value >= UpLevelRes.Value)
            {
                acEd.WriteMessage("\n ОШИБКА!!! Отметка низа стен больше или равны отметке верха стен. Работа программы прекращена!!!");
                return;
            }




            // старт транзакции
            using (Db.Transaction acTrans = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {

                Db.TypedValue[] acTypValAr = new Db.TypedValue[1];
                acTypValAr.SetValue(new Db.TypedValue((int)Db.DxfCode.Start, "INSERT"), 0);
                Ed.SelectionFilter acSelFtr = new Ed.SelectionFilter(acTypValAr);


                Ed.PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection(acSelFtr);
                if (acSSPrompt.Status == Ed.PromptStatus.OK)
                {

                    Ed.SelectionSet acSSet = acSSPrompt.Value;
                    foreach (Ed.SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            Db.Entity acEnt = acTrans.GetObject(acSSObj.ObjectId,
                                                     Db.OpenMode.ForRead) as Db.Entity;
                            if (acEnt != null)
                            {
                                if (acEnt is Db.BlockReference)
                                {
                                    Db.BlockReference acBlRef = (Db.BlockReference)acEnt;
                                    Db.BlockTableRecord blr = (Db.BlockTableRecord)acTrans.GetObject(acBlRef.DynamicBlockTableRecord,
                                                                                                    Db.OpenMode.ForRead);
                                    Db.BlockTableRecord blr_nam = (Db.BlockTableRecord)acTrans.GetObject(blr.ObjectId,
                                                                                                Db.OpenMode.ForRead);
                                    // тут лежит имя блока, в том числе динамческого блока
                                    String acBlock_nam = blr_nam.Name.ToUpper();


                                    if (acBlock_nam.Trim().Contains("Отв с мар".ToUpper()) |
                                        (acBlock_nam.Trim().Contains("NSC_OpeningInWallMarker_v".ToUpper())))
                                    {
                                        Opening op = new Opening();


                                        //Точка центра блока
                                        op.pCenter = new Gem.Point3d(
                                                            (acBlRef.GeometricExtents.MaxPoint.X - acBlRef.GeometricExtents.MinPoint.X) / 2 + acBlRef.GeometricExtents.MinPoint.X,
                                                            (acBlRef.GeometricExtents.MaxPoint.Y - acBlRef.GeometricExtents.MinPoint.Y) / 2 + acBlRef.GeometricExtents.MinPoint.Y,
                                                            0);
                                        //Размер блока
                                        op.MarkerRadius = (acBlRef.GeometricExtents.MaxPoint.X - acBlRef.GeometricExtents.MinPoint.X) / 2;

                                        //точка вставки блока
                                        op.insertPoint = acBlRef.Position;


                                        //Зана проверки самого блока
                                        //########################################################################################
                                        //проверяем на зеркальность
                                        op.ScaleFactor = acBlRef.ScaleFactors;

                                        //проверяем на поворот
                                        op.Rotation = acBlRef.Rotation;
                                        //########################################################################################
                                        //Конец заны проверки самого блока



                                        if (blr.HasAttributeDefinitions)
                                        {
                                            Db.AttributeCollection attrCol = acBlRef.AttributeCollection;
                                            if (attrCol.Count > 0)
                                            {
                                                foreach (Db.ObjectId AttID in attrCol)
                                                {
                                                    Db.AttributeReference acAttRef = acTrans.GetObject(AttID,
                                                                            Db.OpenMode.ForRead) as Db.AttributeReference;

                                                    if (acAttRef.Tag == "ОТМ_НИЗА")
                                                        op.otm_n = acAttRef.TextString;

                                                    if (acAttRef.Tag == "ВЫСОТА")
                                                        op.visota = acAttRef.TextString;
                                                }
                                            }   //Проверка что кол аттрибутов больше 0
                                        }  //Проверка наличия атрибутов



                                        // Тут еще нужно считать динамический параметр "Ширина"
                                        //и проверить отверстие на квадратность
                                        Db.DynamicBlockReferencePropertyCollection acBlockDynProp = acBlRef.DynamicBlockReferencePropertyCollection;
                                        if (acBlockDynProp != null)
                                        {
                                            foreach (Db.DynamicBlockReferenceProperty obj in acBlockDynProp)
                                            {
                                                if (obj.PropertyName == "Ширина")
                                                    op.shirina = Double.Parse(obj.Value.ToString());

                                                if (obj.PropertyName == "Глубина")
                                                    op.glubina = Double.Parse(obj.Value.ToString());
                                            }
                                        }


                                        openingList.Add(op);
                                    }  //Проверка имени блока
                                }   //Проверка, что объект это ссылка на блок
                            }
                        }
                    }
                }








                outError.Add("BlockCounts", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 100, ErrorName = "Количество проверенных блоков" });//Количество проверенных блоков 
                outError.Add("ScaleFactors", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 101, ErrorName = "-проверка на нарушение масштаба блока" });//проверяем на зеркальность 
                outError.Add("Rotation", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 102, ErrorName = "-проверка на нарушением угла поворота блока" }); //проверяем на поворот 


                outError.Add("outOfGeometricExtents", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 103, ErrorName = "-проверка на попадание в заданный дипазон отметок этажа" });//проверка на попадание в заданный дипазон отметок

                outError.Add("attrError_10", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 10, ErrorName = "-проверка атрибута ОТМ_НИЗА на посторонние записи" });
                outError.Add("attrError_11", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 11, ErrorName = "-проверка атрибута ОТМ_НИЗА на попадание в ДОПУСТИМЫЙ диапазон -15 ... +150 метров " }); // 


                outError.Add("attrError_20", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 20, ErrorName = "-проверка атрибута ВЫСОТА на посторонние записи" }); //проверка атрибута ВЫСОТА на посторонние записи 
                outError.Add("attrError_21", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 21, ErrorName = "-проверка атрибута ВЫСОТА отрицательное значение" }); //проверка атрибута ВЫСОТА отрицательное значение
                outError.Add("attrError_22", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 22, ErrorName = "-проверка атрибута ВЫСОТА на черезмерную точность" }); //проверка атрибута ВЫСОТА отрицательное значение

                outError.Add("DynPropError_50", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 50, ErrorName = "-проверка динамического свойства ШИРИНА на черезмерную точность " }); //

                outError.Add("DynPropError_60", new ErroeMetric() { ErrorCount = 0, ErrorColorIndex = 60, ErrorName = "-проверка динамического свойства ГЛУБИНА на черезмерную точность " }); //

                
                //TODO блок рекомендций
                //Хотелось бы, что бы программа рекомендовала, какие отверстия немного не доходят до пола или потолка и их можно было бы несколько увеличить
                //для улучшения технологичности конструкции.


                //Блок проверок
                foreach (Opening i in openingList)
                {

                    string n = "BlockCounts";
                    outError[n].ErrorCount++;

                    //проверяем на зеркальность
                    if (i.ScaleFactor.X != i.ScaleFactor.Y)
                    {
                        n = "ScaleFactors";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }
                    //проверяем на поворот
                    if (i.Rotation != 0)
                    {
                        n = "Rotation";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }

                    //проверка на попадание в заданный дипазон отметок этажа 
                    if ((i.Otm_Niza < DownLevelRes.Value) | (i.Otm_v > UpLevelRes.Value))
                    {
                        n = "outOfGeometricExtents";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }


                    //проверка атрибута ОТМ_НИЗА на посторонние записи
                    if (IsNotNumber(i.otm_n) | i.otm_n.Contains(","))
                    {
                        n = "attrError_10";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }

                    //проверка атрибута ОТМ_НИЗА на попадание в ДОПУСТИМЫЙ диапазон -15 ... +150 метров
                    if (i.Otm_Niza < -15 && i.Otm_Niza > 150)
                    {
                        n = "attrError_11";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }

                    //проверка атрибута ВЫСОТА на посторонние записи 
                    if (IsNotNumber(i.visota) | i.visota.Contains(","))
                    {
                        n = "attrError_20";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }

                    //проверка атрибута ВЫСОТА отрицательное значение
                    if (i.OpeningHight < 0)
                    {
                        n = "attrError_21";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }

                    //проверка на черезмерную точность динамического свойства ШИРИНА
                    if (i.shirina != Math.Truncate(i.shirina))
                    {
                        n = "DynPropError_50";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }

                    //проверка на черезмерную точность динамического свойства ГЛУБИНА
                    if (i.glubina != Math.Truncate(i.glubina))
                    {
                        n = "DynPropError_60";
                        outError[n].ErrorCount++;
                        ErrorMarker(acCurDb, acTrans, i.pCenter, i.MarkerRadius + outError[n].ErrorColorIndex, outError[n].ErrorColorIndex);
                    }



                }// End Блок проверок
                acTrans.Commit();
            }



            //Вывод статистики проверки
            acEd.WriteMessage("\n Command: BX_OPENING_TESTLEVEL" +
                              "\n Application version 1.0.6199.33632" +
                              "\n" + DownLevelOpt.Message.ToString() + DownLevelRes.Value.ToString()+
                              "\n" + UpLevelOpt.Message.ToString() + UpLevelRes.Value.ToString()+
                              "\n Общее количество блоков и Количество блоков не прошедших проверки:"
                              );


            foreach (KeyValuePair<string, ErroeMetric> i in outError)
                acEd.WriteMessage("\n" + i.Value.ErrorName + " (цвет маркера - " + i.Value.ErrorColorIndex + ") : " + i.Value.ErrorCount + ";");

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p">координаты центра окружности</param>
        /// <param name="r">радиус окружности</param>
        private static void ErrorMarker(Db.Database acCurDb, Db.Transaction acTrans, Gem.Point3d p, double r, short color)
        {


            // Открытие таблицы Блоков для чтения
            Db.BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Db.OpenMode.ForRead) as Db.BlockTable;

            // Открытие записи таблицы Блоков пространства Модели для записи
            Db.BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[Db.BlockTableRecord.ModelSpace],
                                                                            Db.OpenMode.ForWrite) as Db.BlockTableRecord;

            // Создание окружности
            Db.Circle acCircle = new Db.Circle();

            acCircle.Center = p;
            acCircle.Radius = r;
            acCircle.ColorIndex = color;
            acCircle.LineWeight = Db.LineWeight.LineWeight070;

            acCircle.SetDatabaseDefaults();
            // Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acCircle);
            acTrans.AddNewlyCreatedDBObject(acCircle, true);

        }

        private static bool IsNotNumber(string input)
        {

            bool ret = false;

            input = input.Replace('+', '.').Replace('-', '.');

            foreach (char c in input)
            {


                if (c != '.')
                {
                    if (!Char.IsNumber(c))
                    {
                        ret = true;
                    }
                }
            }
            return ret;
        }
    }
}
