using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ExcelDataReader;
using OpenXmlPowerTools;

/// <summary>
/// Генератор Артурыч.
/// Позволяет генерировать документы из эсельки xlsx и шаблонов docx.
/// 
/// Artur Krivtsov (gwartur) <gwartur@gmail.com> | Made in Russia
/// Artur Krivtsov © 2020
/// https://wartur.ru
/// BSD 3-Clause License
/// 
/// Вообще я пишу на PHP, последний раз был С# году в 2009, поэтому может быть местами говнокод
/// </summary>
namespace CsvToDocx
{
    
    class Program
    {
        const string excelFileName = @"data\input.xlsx";
        const string templateDirectory = @"data\templates";
        const string resultDirectory = @"data\result";

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            Console.WriteLine("Здравствуйте я генератор Артурыч. Я буду помогать вам создавать бумажки!");
            Console.WriteLine("==================");

            // загружаем Excel файл
            if (File.Exists(excelFileName))
            {
                DirectoryInfo dirInfo;
                FileInfo[] templateFiles;

                Console.WriteLine("Читаем директорию с шаблонами: " + templateDirectory);
                if (Directory.Exists(templateDirectory))
                {
                    // загружаем файлы шаблона
                    dirInfo = new DirectoryInfo(templateDirectory);
                    templateFiles = dirInfo.GetFiles("*.docx"); // работаем только с docx
                    Console.WriteLine("Директория прочитана");
                }
                else
                {
                    Console.WriteLine("Директории не найдено. Создайте её и наполните файлами Word ('.docx'). Из них мы буедм создавать другие файлы. ФАЙЛЫ 'docX' - новые файлы, а не '.doc'");
                    Console.WriteLine("Нажмите любую клавишу");
                    Console.ReadKey();
                    return;
                }


                // в директории result создаем резульат F_I_O_docName.docx
                Console.WriteLine("Читаем директорию с результатом: " + resultDirectory);
                DirectoryInfo diResult;
                if (Directory.Exists(resultDirectory))
                {
                    Console.WriteLine("Директория обнаружена. Чистим её");

                    // если есть данные, чистим всех их
                    diResult = new DirectoryInfo(resultDirectory);
                    foreach (FileInfo file in diResult.GetFiles())
                    {
                        file.Delete();
                    }
                }
                else
                {
                    Console.WriteLine("Директории не обнаружено. Создаю директорию: " + resultDirectory);

                    // если нет директории, создадим её
                    Directory.CreateDirectory(resultDirectory);
                }

                // открываем экселку
                Console.WriteLine("Открываем Excel");
                var streamExcel = File.Open(excelFileName, FileMode.Open, FileAccess.Read);
                IExcelDataReader readerExcel = ExcelReaderFactory.CreateReader(streamExcel);
                Console.WriteLine("Смог открыть Excel файл " + excelFileName + ". Читаем заголовки:");

                // прочитаем заголовок
                readerExcel.Read();
                Dictionary<int, string> headers = new Dictionary<int, string>();
                for (int key = 0; key < readerExcel.FieldCount; ++key)
                {
                    string valueString = readerExcel.GetString(key);
                    if(!String.IsNullOrEmpty(valueString))
                    {
                        string headerColumn = readerExcel.GetString(key);
                        headers.Add(key, headerColumn);
                        Console.WriteLine("Обнаружил столбец: " + headerColumn);
                    }
                }

                // читаем excel файл и генерируем файлы
                Console.WriteLine("");
                Console.WriteLine("Начинаем генерировать файлы");
                int currentLine = 1;
                do
                {
                    currentLine++;
                    while (readerExcel.Read())
                    {
                        // название файла
                        string fileName = "";
                        // получить первые 3 столбца
                        foreach(KeyValuePair<int, string> keyValie in headers)
                        {
                            if (keyValie.Key > 2) break;
                            fileName += readerExcel.GetString(keyValie.Key);
                        }
                        if(fileName == "")
                        {
                            Console.WriteLine("В первых 3 столбцах не было данных. Генерируем название файла резервным вариантом - номером строки");
                            fileName = currentLine.ToString();
                        }

                        Console.WriteLine(" Обработка данных для: " + fileName);
                        foreach (FileInfo template in templateFiles)
                        {
                            // копируем файл из шаблона
                            string newResultFilename = resultDirectory + @"\" + fileName + "_" + template.Name;
                            File.Copy(template.FullName, resultDirectory + @"\" + fileName + "_" + template.Name);

                            // открываем докфайл
                            WordprocessingDocument wordDoc = WordprocessingDocument.Open(newResultFilename, true);

                            // чистим код
                            SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                            {
                                //NormalizeXml = true, // докумнт начинает падать
                                //Additional settings if required
                                //AcceptRevisions = true,
                                //RemoveBookmarks = true,
                                //RemoveComments = true,
                                //RemoveGoBackBookmark = true,
                                //RemoveWebHidden = true,
                                //RemoveContentControls = true,
                                //RemoveEndAndFootNotes = true,
                                //RemoveFieldCodes = true,
                                //RemoveLastRenderedPageBreak = true,
                                //RemovePermissions = true,
                                RemoveProof = true,
                                RemoveRsidInfo = true,
                                RemoveSmartTags = true,
                                //RemoveSoftHyphens = true,
                                //ReplaceTabsWithSpaces = true,
                            };
                            MarkupSimplifier.SimplifyMarkup(wordDoc, settings);

                            string docText = null;
                            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                            {
                                docText = sr.ReadToEnd();
                            }
                            // docText = docText.Replace("</w:t></w:r><w:r><w:t>", ""); // убираем лишние теги (еще читим код) ... не уверен в конструкции... могут быть глюки
                            foreach (KeyValuePair<int, string> keyValie in headers)
                            {
                                string replaceFrom = keyValie.Value;
                                object replaceToObject = readerExcel.GetValue(keyValie.Key);
                                string replaceTo = (replaceToObject == null) ? "" : replaceToObject.ToString();
                                docText = docText.Replace("%" + replaceFrom + "%", replaceTo); // заменяем в файле согласно заголовкам
                            }

                            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                            {
                                sw.Write(docText);
                            }

                            wordDoc.Close();

                            Console.WriteLine("  => Файл " + newResultFilename + " упешно создан и обработан");
                        }
                    }
                } while (readerExcel.NextResult());
                Console.WriteLine("Генерация выполнена. Результат находится в папке 'result'. Не забывайте проверять перед подписью. Удачи!");
            }
            else
            {
                Console.WriteLine("Файл 'input.xlsx' не обнаружен. Добавьте файл с данными. В первой строке должны быть названия столбцов. Далее данные. В шаблонах вы должны указать плейсхолдеры %Название%");
            }

            Console.WriteLine("Нажмите любую клавишу");
            Console.ReadKey();

        }
    }
}
