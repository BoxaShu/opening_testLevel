//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace Opening_testLevel
//{
//    class sec
//    {
//    }
//}


//using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

using System.Management;

using System.IO;
using System.Net;
using System.Text;

using App = Autodesk.AutoCAD.ApplicationServices;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Gem = Autodesk.AutoCAD.Geometry;
using Rtm = Autodesk.AutoCAD.Runtime;


namespace Opening_testLevel
{
    static class sec
    {
        public static int coat;

        private const string CrLf = "\n";
        public static int CheckVER(string url, string ver, string command)
        {
            //Функция получения ответа с сервера о легальности и правильности версии программы
            //' Получениеn текущего документа и базы данных
            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            Db.Database acCurDb = acDoc.Database;
            Ed.Editor acEd = acDoc.Editor;

            string MachineName = Environment.MachineName.ToUpper().ToString();
            string UserName = Environment.UserName.ToUpper().ToString();
            //string DllName = My.Application.Info.AssemblyName.ToUpper.ToString;
            string DllName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString();
            int ret = 0;

            try
            {
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                myHttpWebRequest.Method = "POST";
                myHttpWebRequest.Referer = "http://yandex.ru";
                myHttpWebRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; InfoPath.2; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022; .NET CLR 1.1.4322)2011-10-16 23:07:51";
                myHttpWebRequest.Accept = "*/*";
                myHttpWebRequest.Headers.Add("Accept-Language", "ru");
                myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";

                //ставим False, чтобы не делать автоматического перенаправления
                myHttpWebRequest.AllowAutoRedirect = false;

                string sQueryString = encrypted(CheckHDSN() + "|" + MachineName + "|" + UserName + "|" + DllName + "|" + command, ver);
                //Тут собственн и храним запрос к скрипту армирования
                sQueryString = "Inp=" + sQueryString + "&Ver=" + ver;
                //00.00.0055"

                byte[] ByteArr = System.Text.Encoding.GetEncoding(1251).GetBytes(sQueryString);
                myHttpWebRequest.ContentLength = ByteArr.Length;
                myHttpWebRequest.GetRequestStream().Write(ByteArr, 0, ByteArr.Length);

                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                //делаем запрос
                myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                StreamReader myStreamReader = new StreamReader(myHttpWebResponse.GetResponseStream(), Encoding.GetEncoding(1251));
                string strData1 = myStreamReader.ReadToEnd();
                //TextBox1.Text = decrypted(strData1, ver)
                string temp_string = decrypted(strData1, ver).ToString().Trim();

                if (temp_string.Substring(0, "everything is correct".Length) == "everything is correct")
                {
                    //If Left(temp_string, "everything is correct".Length) = "everything is correct" Then
                    acDoc.Editor.WriteMessage(CrLf + temp_string);
                    ret = 1;
                    return ret;
                }
                else
                {
                    acDoc.Editor.WriteMessage(CrLf + temp_string);
                    return ret;
                }

                //Проверка новой версии
                //Наша программа будет проверять обновление сравнивая своя версию и версия на сайте.
                //Код кнопки:
                //Dim assemblyVersion As String = My.Application.Info.Version.ToString
                //Dim webClient As New Net.WebClient() With {.Proxy = Nothing}
                //If webClient.DownloadString("http://localhost/Version.php") <> assemblyVersion Then
                //MsgBox("Доступна новая версия программы!")
                //Else
                //MsgBox("У вас последняя версия программы!")
                //End If
                //
                //php код:
                //<?php
                // echo "1.0.0.0";
                //?>
                //ВАЖНО: на странице "http://localhost/Version.php" не должно быть рекламы и прочей нечести!			
                //Скачать новую версию можно так:
                //My.Computer.Network.DownloadFile("www.test.ru\1.txt", "c:\1.txt")
                //


            }
            catch (WebException ex)
            {
                //MessageBox.Show("Проверка не пройдена: " & ex.Message)
                acDoc.Editor.WriteMessage(CrLf + "нет доступа к сети интернет:  " + CrLf + ex.Message);
            }

            return ret;
        }

        private static string CheckHDSN()
        {
            //Функция получения номера тома HDD
            //' Получениеn текущего документа и базы данных
            App.Document acDoc = App.Application.DocumentManager.MdiActiveDocument;
            Db.Database acCurDb = acDoc.Database;
            Ed.Editor acEd = acDoc.Editor;

            try
            {
                ManagementObjectSearcher Searcher_L = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_LogicalDisk WHERE DeviceID = 'C:'");
                foreach (ManagementObject queryObj in Searcher_L.Get())
                {
                    queryObj.Get();
                    return queryObj["VolumeSerialNumber"].ToString().Trim();
                    //Exit Function
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("An error occurred while querying for WMI data: VolumeSerialNumber " & ex.Message)
                acDoc.Editor.WriteMessage(CrLf + "Проверка не пройдена: " + ex.Message);

            }
            return "00000000";


        }


        private static string encrypted(string data, string key)
        {
            string[] encrypt = new string[data.Length];
            //Double
            //Кодирование
            int j = 0;
            for (int i = 0; i <= data.Length - 1; i++)
            {
                // encrypt(i) = (Asc(data.Substring(i, 1)) Xor Asc(key.Substring(j, 1))).ToString
                //encrypt(i) = Convert.ToString(Int32.Parse(data.Substring(i, 1)) Xor Int32.Parse(key.Substring(j, 1)))
                string str1 = data.Substring(i, 1);
                string str2 = key.Substring(j, 1);

                char chr1 = Convert.ToChar(str1);
                char chr2 = Convert.ToChar(str2);

                int int1 = Convert.ToInt32(chr1);
                int int2 = Convert.ToInt32(chr2);


                string str = Convert.ToString(int1 ^ int2);

                encrypt[i] = str;
                if (j >= key.Length - 1)
                {
                    j = 0;
                }
                else
                {
                    j = j + 1;
                }
            }
            //Return Join(encrypt, "-")
            return string.Join("-", encrypt);

        }

        private static string decrypted(string data, string key)
        {

            //Dim temp() As String = Split(data, "-")
            string[] temp = data.Split(Convert.ToChar("-"));
            int j = 0;
            string[] decrypt = new string[temp.Length];
            coat = (decrypt.Length * 0) + 21;
            for (int i = 0; i <= temp.Length - 1; i++)
            {
                //decrypt(i) = Chr(temp(i) Xor Asc(key.Substring(j, 1)))
                //decrypt(i) = (Convert.ToChar(Convert.ToInt32(temp(i)) Xor Convert.ToInt32(key.Substring(j, 1)))).ToString

                string str1 = temp[i];
                string str2 = key.Substring(j, 1);

                //Dim chr1 As Char = Int32.Parse(str1)
                char chr2 = Convert.ToChar(str2);

                int int1 = Int32.Parse(str1);
                int int2 = Convert.ToInt32(chr2);

                int int3 =(int1 ^ int2);
                decrypt[i] = Encoding.Default.GetString(new byte[] { (byte)int3 });


                if (j >= key.Length - 1)
                {
                    j = 0;
                }
                else
                {
                    j = j + 1;
                }
            }


            String str_out = string.Join("", decrypt);
            //Return Join(decrypt, "")
            return str_out;

        }






    }
}