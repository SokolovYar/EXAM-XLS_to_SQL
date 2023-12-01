using OfficeOpenXml;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using static OfficeOpenXml.ExcelErrorValue;
using System.Diagnostics;

ExcelReader Excel = new ExcelReader();
Excel.ExcelToSQL();
Excel.RecordToSQL("testStep", "Просто тест");



public class ExcelReader 
{
    private FileInfo _file;
    private string filePath = "ENG_RUS.xlsx";

    public void ExcelToSQL()
    {
        _file = new FileInfo(filePath);
        if (_file.Exists)
        {
            ExcelPackage package = new ExcelPackage(_file);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            string? origin;
            string? trans;

            using (DBWriter DB = new DBWriter())
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    origin = worksheet.Cells[row, 1].Value?.ToString();
                    trans = worksheet.Cells[row, 3].Value?.ToString();
                    DB.Write(origin, trans);
                }
                package.Dispose();
            }
        }
    }

    public void RecordToSQL(string? word, string? translation, string language = "RUS")
    {
        using (DBWriter DB = new DBWriter())
        {
            DB.Write(word, translation, language);
        }
    }
}

public class DBWriter : IDisposable
{
    private string _connectionPath = "Server=localhost;Database=Dictionary;Trusted_Connection=True;Encrypt=False;";
    private SqlConnection _connection;
    public DBWriter()
        {

        //проверка наличия нужной базы данных
        this.CheckDB();
        _connection = new SqlConnection(_connectionPath);
        _connection.Open();
        }

    private void CheckDB()
    {
        _connection = new SqlConnection("Server=localhost;Database=Master;Trusted_Connection=True;Encrypt=False;");
        _connection.Open();
        SqlCommand command = new SqlCommand($"SELECT database_id FROM sys.databases WHERE Name = 'Dictionary'", _connection);
        SqlDataReader reader = command.ExecuteReader();
        reader.Read();
        if (!reader.HasRows)
        {
            Console.WriteLine("'Dictionary' database has not exist!");
            _connection.Close();
            this.CreateTemplateDB();
        }
        _connection.Close();
    }

    private void CreateTemplateDB()
    {
        //создание новой БД
        _connection = new SqlConnection("Server=localhost;Database=Master;Trusted_Connection=True;Encrypt=False;");
        _connection.Open();
        SqlCommand command = new SqlCommand($"CREATE DATABASE Dictionary", _connection);
        command.ExecuteNonQuery();
        command = new SqlCommand("USE Dictionary",_connection);
        command.ExecuteNonQuery();
        command = new SqlCommand("CREATE TABLE Languages ( ID INT PRIMARY KEY IDENTITY(1,1), LanguageName NVARCHAR(100) NOT NULL);", _connection);
        command.ExecuteNonQuery();
        command = new SqlCommand("CREATE TABLE Words ( ID INT PRIMARY KEY IDENTITY(1,1), Word NVARCHAR(100) NOT NULL);", _connection);
        command.ExecuteNonQuery();
        command = new SqlCommand("CREATE TABLE Translations (ID INT PRIMARY KEY IDENTITY(1,1), Translation NVARCHAR(100) NOT NULL, WordID INT, LanguageID INT, FOREIGN KEY (WordID) REFERENCES Words(ID), FOREIGN KEY (LanguageID) REFERENCES Languages(ID));", _connection);
        command.ExecuteNonQuery();
        command = new SqlCommand("INSERT INTO Languages (LanguageName) VALUES ('RUS'), ('ENG');", _connection);
        command.ExecuteNonQuery();
        _connection.Close();
        Console.WriteLine("New database has been created");
    }

    public void Write(string? word, string? translation, string language = "RUS")
    {
        //проверка есть ли уже слово в словаре
        SqlCommand command = new SqlCommand($"SELECT Word FROM Words WHERE Word = '{word}'", _connection);
        SqlDataReader reader = command.ExecuteReader();

        //если слова нет, то 
        if (!reader.HasRows)
        {
            //добавляем слово в таблицу WORD
            reader.Close();
            command = new SqlCommand($"INSERT INTO Words(Word) VALUES ('{word}');", _connection);
            command.ExecuteNonQuery();
            //Считываем добавленное или имеющееся слово и определяем его ID
            command = new SqlCommand($"SELECT ID FROM Words WHERE Word = '{word}'", _connection);
            reader = command.ExecuteReader();

            reader.Read(); // Проверяем, есть ли результаты
            int WordID = reader.GetInt32(0);

            //добавляем перевод к слову в таблицу Translations
            int languageId = -1;
            if (language == "RUS") languageId = 1;
            reader.Close();

            string[] SplitedTransl = this.SplitToWords(translation);

            foreach (string str in SplitedTransl)
            {
                command = new SqlCommand($"INSERT INTO Translations(Translation, WordID, LanguageID) VALUES ('{str}',{WordID}, {languageId});", _connection);
                command.ExecuteNonQuery();
            }
            
            return;
        }
        Console.WriteLine($"Word '{word}'\t\tis already exist in the DB");
        reader.Close();
        command.Dispose();
        return;
    }

    private string[] SplitToWords(string translation)
    {
        translation = translation.Trim();
        string[] temp = translation.Split(',');

        //удаление лишних пробелов в словах
        for (int i = 0; i < temp.Length; i++)
            temp[i] = temp[i].Trim();
        return temp;
    }

    public void Dispose()
    {
        _connection.Close();
    }
}

