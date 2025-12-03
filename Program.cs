using System;

namespace FootballApp
{
    class Program
    {
        static void Main()
        {
            const string file = "la5.xlsx"; 
            var db = new Database(file);

            try { db.Load(); }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка загрузки: " + ex.Message);
                Console.ReadKey();
                return;
            }

            while (true)
            {
                Console.Clear();
                Console.WriteLine("ФУТБОЛЬНЫЕ КЛУБЫ");
                Console.WriteLine("1) Просмотр стран");
                Console.WriteLine("2) Просмотр клубов");
                Console.WriteLine("3) Просмотр достижений");
                Console.WriteLine("4) Удалить страну");
                Console.WriteLine("5) Удалить клуб");
                Console.WriteLine("6) Добавить страну");
                Console.WriteLine("7) Добавить клуб");
                Console.WriteLine("8) Добавить/обновить достижения клуба");
                Console.WriteLine();
                Console.WriteLine("9) LINQ Запрос: Клубы с победами в ЛЧ");
                Console.WriteLine("10) LINQ Запрос: Кол-во клубов в стране");
                Console.WriteLine("11) LINQ Запрос: Клубы с max победами в ЛЕ");
                Console.WriteLine("12) LINQ Запрос: Страна с max золотом без кубков");
                Console.WriteLine();
                Console.WriteLine("0. ВЫЙТИ ОБНОВЛЕНИЕ ФАЙЛА (с сохранением)");
                Console.Write("\nВыбор: ");

                string choice = Console.ReadLine();

                try
                {
                    switch (choice)
                    {
                        case "1": db.ViewCountries(); break;
                        case "2": db.ViewClubs(); break;
                        case "3": db.ViewAchievements(); break;

                        case "4":
                            Console.Write("ID СТРАНЫ: "); int cid = int.Parse(Console.ReadLine()!);
                            Console.WriteLine(db.DeleteCountry(cid) ? "Удалено" : "Не найдено");
                            break;

                        case "5":
                            Console.Write("ID КЛУБА: "); int clid = int.Parse(Console.ReadLine()!);
                            Console.WriteLine(db.DeleteClub(clid) ? "Удалено" : "Не найдено");
                            break;

                        case "6":
                            Console.Write("НАЗВАНИЕ: "); db.AddCountry(Console.ReadLine()!.Trim());
                            Console.WriteLine("ДОБАВИЛИ В ТАБЛИЦУ");
                            break;

                        case "7":
                            Console.Write("НАЗВАНИЕ КЛУБА: "); string name = Console.ReadLine()!.Trim();
                            Console.Write("ID СТРАНы: "); int countryId = int.Parse(Console.ReadLine()!);
                            db.AddClub(name, countryId);
                            Console.WriteLine("ДОБАВИЛИ В ТАБЛИЦУ");
                            break;

                        case "8":
                            Console.Write("Club ID: "); int clubId = int.Parse(Console.ReadLine()!);
                            Console.WriteLine("Введите достижения (0 если нет):");
                            int[] vals = new int[13];
                            string[] prompts = { "Золото (Z)", "Серебро (S)", "Бронза (B)", "Кубки (K)", "Финалы кубков (FK)",
                                               "ЛЧ выиграно", "ЛЧ проиграно", "ЛЕ выиграно", "ЛЕ проиграно",
                                               "КОК выиграно", "КОК проиграно", "ЛК выиграно", "ЛК проиграно" };
                            for (int i = 0; i < 13; i++)
                            {
                                Console.Write($"{prompts[i]}: ");
                                vals[i] = int.Parse(Console.ReadLine() ?? "0");
                            }
                            db.AddOrUpdateAchievement(clubId,
                                vals[0], vals[1], vals[2], vals[3], vals[4],
                                vals[5], vals[6], vals[7], vals[8], vals[9], vals[10], vals[11], vals[12]);
                            Console.WriteLine("Сохранено");
                            break;

                        case "9":
                            Console.WriteLine("КЛУБЫ,ВЫИГРЫВАВШИЕ ЛЧ:");
                            foreach (int id in db.ClubsWithLCH()) Console.WriteLine($"  Club ID: {id}");
                            break;

                        case "10":
                            Console.Write("Название страны: "); string cn = Console.ReadLine()!;
                            Console.WriteLine($"Клубов: {db.ClubsInCountry(cn)}");
                            break;

                        case "11":
                            Console.WriteLine("МАКСИМУМ ПОБЕД В ЛЕ:");
                            foreach (string s in db.ClubsWithMaxLeagueOfEurope()) Console.WriteLine("  " + s);
                            break;

                        case "12":
                            Console.WriteLine(db.CountryWithMostGoldNoCups());
                            break;

                        case "0":
                            db.Save();
                            Console.WriteLine("Изменения сохранены. До свидания!");
                            return;

                        default:
                            Console.WriteLine("НЕВЕРНЫЙ КЕЙС");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ОШИБКА: " + ex.Message);
                }

                Console.WriteLine("\nНАЖМИТЕ ЛЮБУЮ КЛАВИШУ...");
                Console.ReadKey();
            }
        }
    }
}