using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace FootballApp
{


    //КЛАССЫ СТРАНА,КЛУБ,ДОСТИЖЕНИЯ   ПЕРЕГРУЗКА ТУСТРИНГ
    public class Country
    {
        public int Id { get; set; }
        public string Name { get; set; }

        public Country(int id, string name)
        {
            Id = id;
            Name = name ?? string.Empty;
        }

        public override string ToString() => $"{Id}. {Name}";
    }

    public class Club
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int CountryId { get; set; }

        public Club(int id, string name, int countryId)
        {
            Id = id;
            Name = name ?? string.Empty;
            CountryId = countryId;
        }

        public override string ToString() => $"{Id}. {Name} (Страна ID: {CountryId})";
    }

    public class Achievement
    {
        public int ClubId { get; set; }
        public int Z { get; set; }   
        public int S { get; set; }   
        public int B { get; set; }   
        public int K { get; set; }   
        public int FK { get; set; }  
        public int LCH { get; set; } 
        public int FLCH { get; set; }
        public int LE { get; set; }
        public int FLE { get; set; }
        public int KOK { get; set; }
        public int FKOK { get; set; }
        public int LK { get; set; }
        public int FLK { get; set; }

        public Achievement(int clubId, int z, int s, int b, int k, int fk,
            int lch, int flch, int le, int fle, int kok, int fkok, int lk, int flk)
        {
            ClubId = clubId;
            Z = z;
            S = s;
            B = b;
            K = k;
            FK = fk;
            LCH = lch;
            FLCH = flch;
            LE = le;
            FLE = fle;
            KOK = kok;
            FKOK = fkok;
            LK = lk;
            FLK = flk;
        }

        public override string ToString()
        {
            return $"ClubId={ClubId} | Ч: {Z} | К: {K} | ЛЧ: {LCH} | ЛЕ: {LE} | КОК: {KOK} | ЛК: {LK}";
        }
    }



    //РЕАЛИЗОВАТЬ РАБОТУ С ДАННЫМИ 3 ТАБЛИЦЫ(ЛИСТА),СОЗДАЁМ 3 СПИСКА ОБЪЕКТОВ,ИСПОЛЬЗУЕМ ВОРКШИТ
    public class Database
    {
        private readonly string _fileName;
        private Workbook _workbook;
        private Worksheet _wsCountries, _wsClubs, _wsAchievements;

        public List<Country> Countries { get; private set; } = new();
        public List<Club> Clubs { get; private set; } = new();
        public List<Achievement> Achievements { get; private set; } = new();

        public Database(string fileName) => _fileName = fileName;




        //СЧИТЫВАЕМ СТРОКИ С ЛИСТОВ ,СОЗДАЁМ ОБЪЕКТЫ И ВОЗВРАЩАЕМ 3 НОВЫХ СПИСКА
        public void Load()
        {
            _workbook = new Workbook(_fileName);
            _wsCountries = _workbook.Worksheets["Страны"];
            _wsClubs = _workbook.Worksheets["Клубы"];
            _wsAchievements = _workbook.Worksheets["Достижения"];

            if (_wsCountries == null || _wsClubs == null || _wsAchievements == null)
                throw new Exception("Один из листов не найден!");

            Countries = ReadCountries();
            Clubs = ReadClubs();
            Achievements = ReadAchievements();
        }


        //ПЕРЕЗАПИСЫВАЕМ ДАННЫЕ И СОХРАНЯЕМ
        public void Save()
        {
            ClearSheet(_wsCountries);
            ClearSheet(_wsClubs);
            ClearSheet(_wsAchievements);

            WriteCountries();
            WriteClubs();
            WriteAchievements();

            _workbook.Save(_fileName);
        }

        private void ClearSheet(Worksheet ws) => ws.Cells.DeleteRows(1, ws.Cells.MaxDataRow);


        //МЕТОДЫ ЧТЕНИЯ 3 ЛИСТОВ 
        private List<Country> ReadCountries()
        {
            var list = new List<Country>();
            var cells = _wsCountries.Cells;
            for (int stroka = 1; stroka <= cells.MaxDataRow; stroka++)
            {
                if (cells[stroka, 0].Value == null) continue;
                list.Add(new Country(cells[stroka, 0].IntValue, cells[stroka, 1].StringValue.Trim()));
            }
            return list;
        }

        private List<Club> ReadClubs()
        {
            var list = new List<Club>();
            var cells = _wsClubs.Cells;
            for (int stroka = 1; stroka <= cells.MaxDataRow; stroka++)
            {
                if (cells[stroka, 0].Value == null) continue;
                list.Add(new Club(
                    cells[stroka, 0].IntValue,
                    cells[stroka, 1].StringValue.Trim(),
                    cells[stroka, 2].IntValue));
            }
            return list;
        }

        private List<Achievement> ReadAchievements()
        {
            var list = new List<Achievement>();
            var cells = _wsAchievements.Cells;
            for (int stroka = 1; stroka <= cells.MaxDataRow; stroka++)
            {
                if (cells[stroka, 0].Value == null) continue;
                list.Add(new Achievement(
                    GetInt(cells[stroka, 0]),
                    GetInt(cells[stroka, 1]),
                    GetInt(cells[stroka, 2]),
                    GetInt(cells[stroka, 3]),
                    GetInt(cells[stroka, 4]),
                    GetInt(cells[stroka, 5]),
                    GetInt(cells[stroka, 6]),
                    GetInt(cells[stroka, 7]),

                    GetInt(cells[stroka, 8]),
                    GetInt(cells[stroka, 9]),
                    GetInt(cells[stroka, 10]),
                    GetInt(cells[stroka, 11]),

                    GetInt(cells[stroka, 12]),
                    GetInt(cells[stroka, 13] )));
            }
            return list;
        }

        //ПУСТЫЕ ЯЧЕЙКИ ДЛЯ 3 ЛИСТА
        private int GetInt(Cell cell) => cell?.Type == CellValueType.IsNumeric ? cell.IntValue : 0;


        //ЗАПИСИ ЯЧЕЕК В ЛИСТ БЕЗ ПЕРВОЙ,ТАК КАК В 0 СТРОКЕ ЗАГОЛОВКИ
        private void WriteCountries()
        {
            for (int i = 0; i < Countries.Count; i++)
            {
                _wsCountries.Cells[i + 1, 0].PutValue(Countries[i].Id);
                _wsCountries.Cells[i + 1, 1].PutValue(Countries[i].Name);
            }
        }

        private void WriteClubs()
        {
            for (int i = 0; i < Clubs.Count; i++)
            {
                var c = Clubs[i];
                _wsClubs.Cells[i + 1, 0].PutValue(c.Id);
                _wsClubs.Cells[i + 1, 1].PutValue(c.Name);
                _wsClubs.Cells[i + 1, 2].PutValue(c.CountryId);
            }
        }

        private void WriteAchievements()
        {
            for (int i = 0; i < Achievements.Count; i++)
            {
                var a = Achievements[i];
                var stroka = i + 1;
                _wsAchievements.Cells[stroka, 0].PutValue(a.ClubId);
                _wsAchievements.Cells[stroka, 1].PutValue(a.Z);
                _wsAchievements.Cells[stroka, 2].PutValue(a.S);
                _wsAchievements.Cells[stroka, 3].PutValue(a.B);
                _wsAchievements.Cells[stroka, 4].PutValue(a.K);
                _wsAchievements.Cells[stroka, 5].PutValue(a.FK);
                _wsAchievements.Cells[stroka, 6].PutValue(a.LCH);
                _wsAchievements.Cells[stroka, 7].PutValue(a.FLCH);
                _wsAchievements.Cells[stroka, 8].PutValue(a.LE);
                _wsAchievements.Cells[stroka, 9].PutValue(a.FLE);
                _wsAchievements.Cells[stroka, 10].PutValue(a.KOK);
                _wsAchievements.Cells[stroka, 11].PutValue(a.FKOK);
                _wsAchievements.Cells[stroka, 12].PutValue(a.LK);
                _wsAchievements.Cells[stroka, 13].PutValue(a.FLK);
            }
        }

        //ПРОСМОТР(ВЫВОД ВСЕХ ОБЪЕКТОВ В КОНСОЛЬ ПЕРЕГРУЗКА TO STRING),УДАЛЕНИЕ
        public void ViewCountries() => Countries.ForEach(c => Console.WriteLine(c));
        public void ViewClubs() => Clubs.ForEach(c => Console.WriteLine(c));
        public void ViewAchievements() => Achievements.ForEach(a => Console.WriteLine(a));

        public bool DeleteCountry(int id) => Countries.RemoveAll(x => x.Id == id) > 0;
        public bool DeleteClub(int id)
        {
            bool removed = Clubs.RemoveAll(x => x.Id == id) > 0;
            Achievements.RemoveAll(a => a.ClubId == id);
            return removed;
        }
        public bool DeleteAchievement(int clubId) => Achievements.RemoveAll(a => a.ClubId == clubId) > 0;


        //ДОБАВИТЬ ЭЛЕМЕНТ:СТРАНУ,КЛУБ ,ДОСТИЖЕНИЕ
        public void AddCountry(string name)
        {
            int id = Countries.Any() ? Countries.Max(c => c.Id) + 1 : 1;
            Countries.Add(new Country(id, name));
        }

        public void AddClub(string name, int countryId)
        {
            if (!Countries.Any(c => c.Id == countryId))
                throw new Exception("НЕТ СТРАНЫ С ТАКИМ АЙДИ!");
            int id = Clubs.Any() ? Clubs.Max(c => c.Id) + 1 : 1;
            Clubs.Add(new Club(id, name, countryId));
        }

        public void AddOrUpdateAchievement(int clubId, int z, int s, int b, int k, int fk,
            int lch, int flch, int le, int fle, int kok, int fkok, int lk, int flk)
        {
            if (!Clubs.Any(c => c.Id == clubId))
                throw new Exception("ФУТБОЛЬНОГО КЛУБА НЕТ В БАЗЕ ДАННЫХ!");
            Achievements.RemoveAll(a => a.ClubId == clubId);
            Achievements.Add(new Achievement(clubId, z, s, b, k, fk, lch, flch, le, fle, kok, fkok, lk, flk));
        }

        //СФОРМИРОВАТЬ 4 LINQ ЗАПРОСА,2 ПЕРЕЧНЯ,2 ЗНАЧЕНИЯ

        //1) ВЫВЕСТИ СПИСОК КОМАНД КОТОРЫЕ ХОТЯ БЫ РАЗ ПОБЕЖДАЛИ В ЛИГЕ ЧЕМПИОНОВ, LCH>0
        public IEnumerable<int> ClubsWithLCH()
        {
            var query =
                from a in Achievements
                where a.LCH > 0
                select a.ClubId;
            return query;
        }

        //2)ЗАПРОС ВЫВОДИТ КОЛИЧЕСТВО ФУТБОЛЬНЫХ КЛУБОВ ОПРЕДЕЛЕННОЙ СТРАНЫ
        public int ClubsInCountry(string countryName)
        {
            var query =
                from club in Clubs
                join country in Countries on club.CountryId equals country.Id
                where country.Name == countryName
                select club;
            return query.Count();
        }

        //3) ПЕРЕЧЕНЬ КОМАНД ВЫИГРЫВАВШИХ ЛИГУ ЕВРОПЫ ЛЕ
        public IEnumerable<string> ClubsWithMaxLeagueOfEurope()
        {
            int max = Achievements.Max(a => a.LE);

            var query =
                from a in Achievements
                join club in Clubs on a.ClubId equals club.Id
                join country in Countries on club.CountryId equals country.Id
                where a.LE == max
                select club.Name + " (" + country.Name + ") — " + max + " побед в Лиге Европы";

            return query;
        }

        //СТРАНА С НАИБОЛЬШИМ КОЛИЧЕСТВОМ ЗОЛОТЫХ МЕДАЛЕЙ БЕЗ КУБКОВ
        public string CountryWithMostGoldNoCups()
        {
            var query = from a in Achievements
                        where a.K == 0 && a.Z > 0
                        join club in Clubs on a.ClubId equals club.Id
                        join country in Countries on club.CountryId equals country.Id
                        orderby a.Z descending
                        select new { country.Id, country.Name, a.Z };

            var top = query.FirstOrDefault();
            return top != null ? $"Страна ID: {top.Id} ({top.Name}), золотых медалей: {top.Z}" : "Нет данных";
        }
    }
}