using System;
using NetOffice;
using NetOffice.MSProjectApi;
using NetOffice.MSProjectApi.Enums;
using MSProject = NetOffice.MSProjectApi;

namespace NetOfficeProjectSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start");
            // 声明一个操作类
            Application app = new();

            try
            {
                #region 创建一个 Project文件 写入任务数据 并保存
                //// 添加一个项目
                //Project newProject = app.Projects.Add();
                //// 添加任务
                //newProject.Tasks.Add("Task Lists");
                //newProject.Tasks.Add("Task 0");
                //newProject.Tasks.Add("Task 1");
                //newProject.Tasks.Add("Task 2");
                //newProject.Tasks.Add("Task 3");
                //newProject.Tasks.Add("Task 4");
                //newProject.Tasks.Add("Task 5");
                //newProject.Tasks.Add("Task 6");
                //newProject.Tasks.Add("Task 7");
                //newProject.Tasks.Add("Task 8");
                //newProject.Tasks.Add("Task 9");
                //newProject.Tasks.Add("Task 10");
                //// 保存 路径、保存形式、是否备份、是否只读
                //newProject.SaveAs(@$"{AppDomain.CurrentDomain.BaseDirectory}\MSProject{DateTime.Now:yyyyMMddHHmmss}.mpp", PjSaveType.pjDoNotSave, false, false);
                //// 释放
                //newProject = null;

                #endregion

                // 打开一个 mpp 文件




            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Err Info {ex}");
            }
            finally
            {
                // 防止内存溢出
                if (null != app)
                {
                    app.Quit(PjSaveType.pjDoNotSave);
                    app.Dispose();
                }
                Console.WriteLine("Done");
                Console.ReadLine();
            }
        }
    }
}
