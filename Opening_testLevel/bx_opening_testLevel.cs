using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using App = Autodesk.AutoCAD.ApplicationServices;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Gem = Autodesk.AutoCAD.Geometry;
using Rtm = Autodesk.AutoCAD.Runtime;

// [assembly: Rtm.CommandClass(typeof(MyClassSerializer.Commands))]


namespace Opening_testLevel
{
    public class Commands
    {
        //Проверка блоков отверстий, в Autocad, на попадание в отметку этажа
        [Rtm.CommandMethod("bx_opening_testLevel")]
        static public void bx_opening_testLevel()
        {
            // Получение текущего документа и базы данных
            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            Db.Database acCurDb = acDoc.Database;
            Ed.Editor acEd = acDoc.Editor;


            Ed.PromptDoubleOptions DownLevelOpt = new Ed.PromptDoubleOptions("\n введи нижнюю отметку стены: ");
            DownLevelOpt.AllowNone = false;
            Ed.PromptDoubleResult DownLevelRes = acEd.GetDouble(DownLevelOpt);
            if (DownLevelRes.Status != Ed.PromptStatus.OK)
            {
                return;
            }

            Ed.PromptDoubleOptions UpLevelOpt = new Ed.PromptDoubleOptions("\n введи верхнюю отметку стены: ");
            UpLevelOpt.AllowNone = false;
            Ed.PromptDoubleResult UpLevelRes = acEd.GetDouble(UpLevelOpt);
            if (UpLevelRes.Status != Ed.PromptStatus.OK)
            {
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

                    // Открытие таблицы Блоков для чтения
                    Db.BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Db.OpenMode.ForRead) as Db.BlockTable;

                    // Открытие записи таблицы Блоков пространства Модели для записи
                    Db.BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[Db.BlockTableRecord.ModelSpace],
                                                                                    Db.OpenMode.ForWrite) as Db.BlockTableRecord;


                    
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
                                    Db.BlockReference acBlRef = (Db.BlockReference) acEnt;
                                    
                                    
                                    // тут нужна проверка имени блока.
                                    //тут выясняю истинное имя блока для последующего обновления атрибутов.
                                    //Проверяю является ли выделенный блок динамическим
                                    //Получаю настоящие/родное имя динамического блока

                                    Db.BlockTableRecord blr = (Db.BlockTableRecord)acTrans.GetObject(acBlRef.DynamicBlockTableRecord, 
                                                                                                    Db.OpenMode.ForRead);


                                    
                                    if(blr.HasAttributeDefinitions)
                                    {

                                        Db.BlockTableRecord blr_nam = (Db.BlockTableRecord)acTrans.GetObject(blr.ObjectId, 
                                                                                                    Db.OpenMode.ForRead);
                                        // тут лежит имя блока, в том числе динамческого блока
                                        String acBlock_nam = blr_nam.Name.ToUpper();


                                        //Теперь вот этот вот фрагмент кода на VB.NEt  надо переписать на С#
                                        /*
                               If acBlock_nam.ToUpper Like "Отв с мар*".ToUpper Or
                                    acBlock_nam.ToUpper = "Otverstie".ToUpper Then
                                    ...
                                End If
                                         */

                                        if (acBlock_nam.Trim().Contains("Отв с мар".ToUpper()) | 
                                            (acBlock_nam.Trim().Contains("Otverstie".ToUpper())))
                                        {

                                      


                                        Db.AttributeCollection attrCol = acBlRef.AttributeCollection;
                                    if (attrCol.Count > 0)
                                    {
                                       
                                        
                                        Double Otm_n = 0;
                                        Double visota = 0;
                                        Double Otm_v = 0;

                                        foreach (Db.ObjectId AttID in attrCol)
                                        {
                                            Db.AttributeReference acAttRef = acTrans.GetObject(AttID,
                                                                    Db.OpenMode.ForRead) as Db.AttributeReference;

                                           

                                            if (acAttRef.Tag == "ОТМ_НИЗА")
                                            {
                                                // Otm_n = Double.Parse(acAttRef.TextString);
                                                //Double.TryParse(acAttRef.TextString, Otm_n);
                                                Double.TryParse(acAttRef.TextString.Replace(',','.'), out Otm_n);

                                            }

                                            if (acAttRef.Tag == "ВЫСОТА")
                                            {
                                                //visota = Double.Parse(acAttRef.TextString);
                                                Double.TryParse(acAttRef.TextString.Replace(',', '.'), out visota);
                                            }


                                           // acEd.WriteMessage("\n " + acAttRef.Tag + " = " + acAttRef.TextString);


                                        }

                                        Otm_v = Otm_n + visota / 1000;

                                        if ((Otm_n < DownLevelRes.Value) | (Otm_v > UpLevelRes.Value))
                                        {

                                            // Создание окружности
                                            Db.Circle acCircle = new Db.Circle();

                                            acCircle.Center = acBlRef.Position;
                                            acCircle.Radius = 111;
                                            acCircle.ColorIndex = 1;
                                            acCircle.LineWeight = Db.LineWeight.LineWeight070;

                                            acCircle.SetDatabaseDefaults();
                                            // Добавление нового объекта в запись таблицы блоков и в транзакцию
                                            acBlkTblRec.AppendEntity(acCircle);
                                            acTrans.AddNewlyCreatedDBObject(acCircle, true);
                                            
                                            //acEd.WriteMessage("\n Косяк");
                                        }
                                        else
                                        {
                                            //acEd.WriteMessage("\n Успешно");
                                        }
                                        


                                        

                                    }   //Проверка что кол аттрибутов больше 0

                                    }  //Проверка имени блока

                                    }  //Проверка наличия атрибутов


                                }   //Проверка, что объект это ссылка на блок
                              //acEnt.ColorIndex = 3;
                            }
                        }
                    }
                }

                acTrans.Commit();

                
            }
        }
    }
}
