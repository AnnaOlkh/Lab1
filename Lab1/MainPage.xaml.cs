using Microsoft.Maui.Controls;
using Microsoft.Maui.Controls.Compatibility;
using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using Grid = Microsoft.Maui.Controls.Grid;
using Microsoft.Maui;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Maui.Storage;
using System.IO;
using System.Linq;
using Lab1.MyGrammar;
namespace Lab1
{
    public partial class MainPage : ContentPage
    {
        private const int start_column_count = 10;
        private const int start_row_count = 20;

        private Entry last_clicked_cell = new Entry();

        private Dictionary<string, string> cell_expressions = new Dictionary<string, string>();

        public MainPage()
        {
            InitializeComponent();
            CreateGrid();
        }
        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }
        private void AddColumnsAndColumnLabels()
        {
            for (int col = 0; col < start_column_count + 1; col++)
            {
                grid.ColumnDefinitions.Add(new Microsoft.Maui.Controls.ColumnDefinition());

                if (col > 0)
                {
                    var label = new Label
                    {
                        Text = GetColumnName(col),
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0);
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                }
            }
        }
        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;

            }
            return columnName;
        }
        private void AddRowsAndCellEntries()
        {
            for (int row = 0; row < start_row_count + 1; row++)
            {
                grid.RowDefinitions.Add(new Microsoft.Maui.Controls.RowDefinition());

                if (row > 0)
                {
                    CreateAndSetRowLabel(row);

                    for (int col = 0; col < start_column_count; col++)
                    {
                        CreateAndSetCell(row, col);
                    }
                }
            }
        }
        private void CreateAndSetRowLabel(int row)
        {
            var label = new Label
            {
                Text = (row).ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, row);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);
        }
        private void CreateAndSetCell(int row, int col)
        {
            var entry = new Entry
            {
                Text = "",
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            entry.TextChanged += EntryTextChanged;
            entry.Focused += EntryFocused;

            string cellName = GetColumnName(col + 1) + row.ToString(); // Correct cell name format
            cell_expressions[cellName] = entry.Text;

            // Store the cell name in the Entry's BindingContext for easy access later
            entry.BindingContext = cellName;

            Grid.SetRow(entry, row);
            Grid.SetColumn(entry, col + 1);
            grid.Children.Add(entry);
        }
        private void EntryFocused(object sender, FocusEventArgs e)
        {
            last_clicked_cell = (Entry)sender; string cellName = (string)last_clicked_cell.BindingContext;
            text_input.Text = cell_expressions[cellName];
        }
        private void EntryTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender == last_clicked_cell)
            {
                text_input.Text = e.NewTextValue;
                string cellName = (string)last_clicked_cell.BindingContext; cell_expressions[cellName] = e.NewTextValue;
            }
        }
        private void MainEntryTextChanged(object sender, TextChangedEventArgs e)
        {
            if (last_clicked_cell != null)
            {
                last_clicked_cell.Text = e.NewTextValue;
                string cellName = (string)last_clicked_cell.BindingContext; cell_expressions[cellName] = e.NewTextValue;
            }
        }
        private void SaveButtonClicked(object sender, EventArgs e)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("GridData");

            int row_count = grid.RowDefinitions.Count;
            int col_count = grid.ColumnDefinitions.Count;

            for (int row = 0; row < row_count + 1; row++)
            {
                for (int col = 0; col < col_count + 1; col++)
                {
                    var entry = grid.Children.FirstOrDefault(e => grid.GetRow(e) == row && grid.GetColumn(e) == col) as Entry;

                    if (entry is not null)
                    {
                        worksheet.Cell(row, col).Value = entry.Text;
                    }
                }
            }

            string folderPath = "D:\\Ann\\k24";
            string file_name = Path.Combine(folderPath, "MyExel.xlsx");

            workbook.SaveAs(file_name);

            DisplayAlert("the file has been saved in project folder", "ok", "ok");
        }
        private void CalculateButtonClicked(object sender, EventArgs e)
        {
            string expr = text_input.Text;
            Evaluator evaluator = new Evaluator(cell_expressions);
            try
            {
                Evaluator.HandleEmptyExpression(expr);
                string cellName = (string)last_clicked_cell.BindingContext;
                var result = evaluator.Evaluate(expr, cellName);
                text_input.Text = result.ToString(); last_clicked_cell.Text = result.ToString();
                // Store the original expression (in case it's a formula)
                cell_expressions[cellName] = expr;
            }
            catch (Exception ex)
            {
                DisplayAlert("Помилка", ex.Message, "OK");
            }
        }
        private async void ExitButtonClicked(object sender, EventArgs e)
        {
            bool answer2 = await DisplayAlert("Підтвердження", "Ви хочете зберегти файл?", "Так", "Ні");
            if (answer2)
            {
                SaveButtonClicked(sender, e);
            }
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                System.Environment.Exit(0);
            }
        }
        private void AddRowButtonClicked(object sender, EventArgs e)
        {
            int new_row = grid.RowDefinitions.Count();


            grid.RowDefinitions.Add(new Microsoft.Maui.Controls.RowDefinition());

            CreateAndSetRowLabel(new_row);

            for (int col = 0; col < grid.ColumnDefinitions.Count; col++)
            {
                CreateAndSetCell(new_row, col);
            }
        }
        private void AddColumnButtonClicked(object sender, EventArgs e)
        {
            int new_col = grid.ColumnDefinitions.Count();

            grid.ColumnDefinitions.Add(new Microsoft.Maui.Controls.ColumnDefinition());

            var label = new Label
            {
                Text = GetColumnName(new_col),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, new_col);
            grid.Children.Add(label);

            for (int row = 1; row < grid.RowDefinitions.Count; row++)
            {
                CreateAndSetCell(row, new_col - 1);
            }
        }
        private void DeleteRowButtonClicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count() > 2)
            {
                int last_row_index = grid.RowDefinitions.Count() - 1;
                grid.RowDefinitions.RemoveAt(last_row_index);

                var elements_to_remove = new List<Microsoft.Maui.IView>();

                foreach (var child in grid.Children)
                {
                    if (grid.GetRow(child) == last_row_index)
                    {
                        elements_to_remove.Add(child);
                    }
                }

                foreach (var element in elements_to_remove)
                {
                    grid.Children.Remove(element);
                }
            }
        }
        private void DeleteColumnButtonClicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count() > 2)
            {
                int last_col_index = grid.ColumnDefinitions.Count() - 1;
                grid.ColumnDefinitions.RemoveAt(last_col_index);

                var elements_to_remove = new List<Microsoft.Maui.IView>();

                foreach (var child in grid.Children)
                {
                    if (grid.GetColumn(child) == last_col_index)
                    {
                        elements_to_remove.Add(child);
                    }
                }

                foreach (var element in elements_to_remove)
                {
                    grid.Children.Remove(element);
                }
            }
        }
        private async void HelpButtonClicked(object sender, EventArgs e)
        {
            string info = @"Лабораторна робота 1.
Ольховська Анна К24
                            
Операції та функції:
    – +, -, *, /, mod, div
    - ^ (pow)
    - mmax(), mmin()
";

            await DisplayAlert("Довідка", info, "OK");
        }
        private async void OpenButtonClicked(object sender, EventArgs e)
        {
            bool answer2 = await DisplayAlert("Підтвердження", "Ви хочете зберегти файл?", "Так", "Ні");
            if (answer2)
            {
                SaveButtonClicked(sender, e);
            }
            string filePath = "D:\\Ann\\k24\\MyExel.xlsx";
            if (File.Exists(filePath))
            {
                LoadExcelFile(filePath); DisplayAlert("Успіх", "Файл успішно відкрито!", "OK");
            }
            else
            {
                DisplayAlert("Помилка", "Файл не знайдено. Спочатку збережіть файл.", "OK");
            }
        }
        private void LoadExcelFile(string filePath)
        {
            try
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(1); // Load the first worksheet 

                // Clear the current grid 
                grid.Children.Clear();
                grid.ColumnDefinitions.Clear();
                grid.RowDefinitions.Clear();

                // Determine the number of rows and columns 
                int rowCount = worksheet.RowsUsed().Count();
                int colCount = worksheet.ColumnsUsed().Count();

                // Rebuild the grid structure 
                for (int col = 0; col <= colCount; col++)
                {
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                }

                for (int row = 0; row <= rowCount; row++)
                {
                    grid.RowDefinitions.Add(new RowDefinition());
                }

                // Populate the grid with data from the Excel file 
                foreach (var cell in worksheet.CellsUsed())
                {
                    int row = cell.Address.RowNumber;
                    int col = cell.Address.ColumnNumber;

                    if (row == 1)
                    {
                        // Add column headers 
                        var label = new Label
                        {
                            Text = cell.GetValue<string>(),
                            VerticalOptions = LayoutOptions.Center,
                            HorizontalOptions = LayoutOptions.Center
                        };
                        Grid.SetRow(label, 0);
                        Grid.SetColumn(label, col);
                        grid.Children.Add(label);
                    }
                    else
                    {
                        // Add cell values 
                        var entry = new Entry
                        {
                            Text = cell.GetValue<string>() ?? string.Empty,
                            VerticalOptions = LayoutOptions.Center,
                            HorizontalOptions = LayoutOptions.Center
                        };
                        entry.TextChanged += EntryTextChanged;
                        entry.Focused += EntryFocused;

                        string cellName = GetColumnName(col) + row;
                        cell_expressions[cellName] = entry.Text;
                        entry.BindingContext = cellName;

                        Grid.SetRow(entry, row);
                        Grid.SetColumn(entry, col);
                        grid.Children.Add(entry);
                    }
                }

                DisplayAlert("Успіх", "Дані завантажено успішно!", "OK");
            }
            catch (Exception ex)
            {
                DisplayAlert("Помилка", $"Не вдалося завантажити файл: {ex.Message}", "OK");
            }
        }

    }

}