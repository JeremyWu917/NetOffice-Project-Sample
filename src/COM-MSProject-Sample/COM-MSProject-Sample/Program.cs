using System;
using System.Data;

namespace COM_MSProject_Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start");
            string fileName = @$"{AppDomain.CurrentDomain.BaseDirectory}\SampleProjectPlanning.mpp";

            // 读取 MSProject 文件
            //DataTable table = new();
            //MSProjectHandle fileProcess = new(fileName, table);
            //Console.WriteLine(fileProcess.LoadMSProject());
            //table = fileProcess.TasksTable;

            //// 编辑 MSProject 文件
            Console.WriteLine("Try to edit MS Project...");
            MSProjectHandle fileProcess = new(fileName);
            Console.WriteLine(fileProcess.EditMSProject("test21", DateTime.Now, DateTime.Now, 3));
            Console.WriteLine("Done");

            Console.ReadLine();
        }
    }
}
