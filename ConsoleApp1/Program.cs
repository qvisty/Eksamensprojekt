
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.IO;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace SchoolManagement
{
    class Student
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public int Grade { get; set; }
        public int Group { get; set; }
        public char Sex { get; set; }
    }

    class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public int Grade { get; set; }
        public int Group { get; set; }
        public char Sex { get; set; }
        public string Education { get; set; }
    }

    class Program
    {
        static List<Employee> LoadEmployees()
        {
            return new List<Employee>
            {
                new Employee { Name = "Alice", Age = 18, Grade = 90 },
                new Employee { Name = "Bob", Age = 19, Grade = 80 },
                new Employee { Name = "Charlie", Age = 17, Grade = 95 },
                new Employee { Name = "David", Age = 18, Grade = 85 },
                new Employee { Name = "Eve", Age = 20, Grade = 75 },
                new Employee { Name = "Frank", Age = 18, Grade = 60 },
                new Employee { Name = "Gina", Age = 19, Grade = 50 },
                new Employee { Name = "Henry", Age = 17, Grade = 55 },
                new Employee { Name = "Ida", Age = 18, Grade = 45 },
                new Employee { Name = "Joan", Age = 20, Grade = 40 }
            };
        }

        static List<Student> LoadStudents()
        {
            
            List<Student> students = new List<Student>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo("C:\\Users\\Jesper\\source\\repos\\ConsoleApp1\\ConsoleApp1\\students.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    Student student = new Student();
                    student.Name = worksheet.Cells[i, 1].Value.ToString();
                    student.Age = int.Parse(worksheet.Cells[i, 2].Value.ToString());
                    student.Grade = int.Parse(worksheet.Cells[i, 3].Value.ToString());
                    // student.Major = worksheet.Cells[i, 4].Value.ToString();
                    students.Add(student);
                }
                package.Dispose();
            }
            return students;
        }


        static void Main(string[] args)
        {
            // Create a list to store students
            var students = new List<Student>();

            // Create a list to store employees and load them
            var employees = new List<Employee>();


            bool firstRun = true;
            // Display the main menu
            while (true)
            {
                // Wait for the user to press a key
                if (!firstRun)
                {
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                };
                firstRun = false;
                // Clear the screen
                Console.Clear();

                // Print the menu
                printMenu();

                // Read the user's choice
                int choice = Convert.ToInt32(Console.ReadLine());

                // Act on the user's choice
                switch (choice)
                {
                    case 1:
                        // List the students
                        ListStudents(students);
                        break;
                    case 2:
                        // Add a student
                        AddStudent(students);
                        break;
                    case 3:
                        // Update a student
                        UpdateStudent(students);
                        break;
                    case 4:
                        // Delete a student
                        DeleteStudent(students);
                        break;
                    case 5:
                        // Search for a student
                        SearchStudent(students);
                        break;
                    case 6:
                        // Show a bar graph of student grades
                        ShowBarGraph(students);
                        break;

                    case 7:
                        // Show a summary of the students
                        ShowSummary(students);
                        break;
                    case 8:
                        // Generate the HTML file
                        GenerateHTML(students);
                        break;
                    case 10:
                        // A. Load demo data for employees
                        employees = LoadEmployees();
                        break;
                    case 11:
                        // B. Load demo data for students
                        students = LoadStudents();
                        break;
                    case 12:
                        // List the students
                        ListEmployees(employees);
                        break;
                    case 99:
                        // Exit the program
                        return;
                    default:
                        Console.WriteLine("Invalid choice.");
                        break;
                }
            }
        }

        private static void printMenu()
        {
            // Show the main menu
            Console.WriteLine("1. List students");
            Console.WriteLine("2. Add a student");
            Console.WriteLine("3. Update a student");
            Console.WriteLine("4. Delete a student");
            Console.WriteLine("5. Search for a student");
            Console.WriteLine("6. Show a bar graph of student grades");
            Console.WriteLine("7. Show a summary of the students");
            Console.WriteLine("8. Generate HTML file");
            Console.WriteLine("----------");
            Console.WriteLine("10. Load demo data for employees");
            Console.WriteLine("11. Load demo data for students");
            Console.WriteLine("12. List employees");
            Console.WriteLine("----------");
            Console.WriteLine("99. Exit");

            Console.Write("Enter your choice: ");
        }

        static void SearchStudent(List<Student> students)
        {
            // Prompt the user for the student's name
            Console.Write("Enter the name of the student to search for: ");
            string name = Console.ReadLine();

            // Find the student in the list
            Student student = students.Find(s => s.Name == name);
            if (student == null)
            {
                Console.WriteLine("Student not found.");
                return;
            }

            // Print the student's information
            Console.WriteLine("Name: " + student.Name);
            Console.WriteLine("Age: " + student.Age);
            Console.WriteLine("Grade: " + student.Grade);
        }

        static void GenerateHTML_old(List<Student> students)
        {
            // Create a new StringBuilder for building the HTML file
            StringBuilder htmlBuilder = new StringBuilder();

            // Write the HTML header
            htmlBuilder.AppendLine("<html>");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine("<title>Student Grades</title>");
            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body>");

            // Write the table of students and their grades
            htmlBuilder.AppendLine("<table>");
            htmlBuilder.AppendLine("<tr>");
            htmlBuilder.AppendLine("<th>Name</th>");
            htmlBuilder.AppendLine("<th>Age</th>");
            htmlBuilder.AppendLine("<th>Grade</th>");
            htmlBuilder.AppendLine("</tr>");
            foreach (var student in students)
            {
                htmlBuilder.AppendLine("<tr>");
                htmlBuilder.AppendLine("<td>" + student.Name + "</td>");
                htmlBuilder.AppendLine("<td>" + student.Age + "</td>");
                htmlBuilder.AppendLine("<td>" + student.Grade + "</td>");
                htmlBuilder.AppendLine("</tr>");
            }
            htmlBuilder.AppendLine("</table>");

            // Write the HTML footer
            htmlBuilder.AppendLine("</body>");
            htmlBuilder.AppendLine("</html>");

            // Write the HTML to a file
            string html = htmlBuilder.ToString();
            File.WriteAllText("students.html", html);
        }

        static void GenerateHTML(List<Student> students)
        {
            // Create a new StringBuilder for building the HTML file
            StringBuilder htmlBuilder = new StringBuilder();

            // Write the HTML header
            htmlBuilder.AppendLine("<html>");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine("<title>Student Grades</title>");

            // htmlBuilder.AppendLine("< script src = 'https://cdn.plot.ly/plotly-latest.min.js' ></ script >");


            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body>");

            // Write the bar chart container
            htmlBuilder.AppendLine("<div id='bar-chart'></div>");

            // Write the JavaScript code for rendering the bar chart
            htmlBuilder.AppendLine("<script>");
            htmlBuilder.AppendLine("var data = [");
            foreach (var student in students)
            {
                htmlBuilder.AppendLine("{ x: '" + student.Name + "', y: " + student.Grade + " },");
            }
            htmlBuilder.AppendLine("];");
            htmlBuilder.AppendLine("var layout = {");
            htmlBuilder.AppendLine("title: 'Student Grades',");
            htmlBuilder.AppendLine("xaxis: { title: 'Student' },");
            htmlBuilder.AppendLine("yaxis: { title: 'Grade' },");
            htmlBuilder.AppendLine("};");
            htmlBuilder.AppendLine("Plotly.newPlot('bar-chart', data, layout, {});");
            htmlBuilder.AppendLine("</script>");

            // Write the HTML footer
            htmlBuilder.AppendLine("</body>");
            htmlBuilder.AppendLine("</html>");

            // Write the HTML to a file
            string html = htmlBuilder.ToString();
            File.WriteAllText("students.html", html);
        }


        static void AddStudent(List<Student> students)
        {
            // Prompt the user for the student's information
            Console.Write("Enter the student's name: ");
            string name = Console.ReadLine();
            Console.Write("Enter the student's age: ");
            int age = Convert.ToInt32(Console.ReadLine());
            Console.Write("Enter the student's grade: ");
            int grade = Convert.ToInt32(Console.ReadLine());

            // Create a new student object
            var student = new Student
            {
                Name = name,
                Age = age,
                Grade = grade
            };

            // Add the student to the list
            students.Add(student);
        }

        static void ShowSummary(List<Student> students)
        {
            // Calculate the average grade
            double total = 0;
            foreach (var student in students)
            {
                total += student.Grade;
            }
            double average = total / students.Count;

            // Find the highest and lowest grades
            int highest = 0;
            int lowest = 100;
            foreach (var student in students)
            {
                if (student.Grade > highest)
                {
                    highest = student.Grade;
                }
                if (student.Grade < lowest)
                {
                    lowest = student.Grade;
                }
            }

            // Write the summary to the Console
            Console.WriteLine($"Number of students: {students.Count}");
            Console.WriteLine($"Average grade: {average:F2}");
            Console.WriteLine($"Highest grade: {highest}");
            Console.WriteLine($"Lowest grade: {lowest}");
        }

        static void ShowBarGraph(List<Student> students)
        {
            // Sort the students by grade in descending order
            students.Sort((s1, s2) => s2.Grade.CompareTo(s1.Grade));

            // Iterate over the students and display a bar graph of their grades
            foreach (var student in students)
            {
                Console.Write(student.Name.PadRight(20));
                Console.Write(student.Grade.ToString().PadRight(6));
                for (int i = 0; i < student.Grade; i += 5)
                {
                    Console.Write("*");
                }
                Console.WriteLine();
            }
        }




        static void ListStudents(List<Student> students)
        {
            // Print the table header
            Console.WriteLine("Students:");
            Console.WriteLine("Name".PadRight(20) + "Age".PadRight(6) + "Grade".PadRight(6));
            Console.WriteLine("-----------------------------------------------");

            // Iterate over the students and print their information
            foreach (var student in students)
            {
                Console.WriteLine(student.Name.PadRight(20) + student.Age.ToString().PadRight(6) + student.Grade.ToString().PadRight(6));
            }
        }
        static void ListEmployees(List<Employee> employees)
        {
            // Print the table header
            Console.WriteLine("Employees:");
            Console.WriteLine("Name".PadRight(20) + "Age".PadRight(6) + "Grade".PadRight(6));
            Console.WriteLine("-----------------------------------------------");

            // Iterate over the employees and print their information
            foreach (var emp in employees)
            {
                Console.WriteLine(emp.Name.PadRight(20) + emp.Age.ToString().PadRight(6) + emp.Grade.ToString().PadRight(6));
            }
        }



        static void UpdateStudent(List<Student> students)
        {
            // Prompt the user for the student's name
            Console.Write("Enter the name of the student to update: ");
            string name = Console.ReadLine();

            // Find the student in the list
            Student student = students.Find(s => s.Name == name);
            if (student == null)
            {
                Console.WriteLine("Student not found.");
                return;
            }

            // Prompt the user for the updated information
            Console.Write("Enter the updated age: ");
            int age = Convert.ToInt32(Console.ReadLine());
            Console.Write("Enter the updated grade: ");
            int grade = Convert.ToInt32(Console.ReadLine());

            // Update the student's information
            student.Age = age;
            student.Grade = grade;

            Console.WriteLine("Student updated.");
        }
        static void DeleteStudent(List<Student> students)
        {
            // Prompt the user for the student's name
            Console.Write("Enter the name of the student to delete: ");
            string name = Console.ReadLine();

            // Find the student in the list
            Student student = students.Find(s => s.Name == name);
            if (student == null)
            {
                Console.WriteLine("Student not found.");
                return;
            }

            // Remove the student from the list
            students.Remove(student);
            Console.WriteLine("Student deleted.");
        }

    }
}




