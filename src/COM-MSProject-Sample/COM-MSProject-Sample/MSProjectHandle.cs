using Microsoft.Office.Interop.MSProject;
using System;
using System.Collections;
using System.Data;

namespace COM_MSProject_Sample
{
    public class MSProjectHandle
    {
        /****************************** 定义私有变量 ********************************/
        private readonly string newLine = Environment.NewLine;
        private readonly object readOnly = true;
        private readonly object readAndWrite = false;
        private readonly object missing = Type.Missing;
        private readonly PjPoolOpen pool = PjPoolOpen.pjPoolReadOnly;
        private readonly ApplicationClass appClass = new();

        /****************************** 定义私有属性 ********************************/
        // MSProject 文件名称
        private string fileName;
        public string FileName
        {
            get { return fileName; }
            set { fileName = value; }
        }
        // 数据表
        private DataTable tasksTable;

        public DataTable TasksTable
        {
            get { return tasksTable; }
            set { tasksTable = value; }
        }

        /****************************** 定义构造函数 ********************************/
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName">MSProject 文件名称</param>
        /// <param name="table">数据表</param>
        public MSProjectHandle(string fileName, DataTable table)
        {
            this.fileName = fileName;
            this.tasksTable = table;
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName">MSProject 文件名称</param>
        public MSProjectHandle(string fileName)
        {
            this.fileName = fileName;
        }

        /****************************** 定义私有属性 ********************************/
        /// <summary>
        /// 读取 MSProject 文件 并写入到 TasksTable
        /// </summary>
        /// <returns></returns>
        public string LoadMSProject()
        {
            ApplicationClass app = null;
            string result;
            string dateStart;
            string dateFinish;
            bool isSummary;
            ArrayList tasks = new();

            try
            {
                app = new();
                app.Visible = false;
                if (app.FileOpen(fileName, readOnly, missing,
                    missing, missing, missing, missing,
                    missing, missing, missing, missing,
                    pool, missing, missing, missing, missing))
                {
                    Tasks taskTable = new(tasksTable);
                    tasksTable = taskTable.CreateTable();
                    Project proj = app.ActiveProject;
                    foreach (Task task in proj.Tasks)
                    {
                        dateStart = task.Start.ToString() == "NA" ? "0000-00-00" : task.Start.ToString();
                        dateFinish = task.Finish.ToString() == "NA" ? "0000-00-00" : task.Finish.ToString();
                        isSummary = task.OutlineChildren.Count != 0;

                        Tasks Tasks = new(task.ID, task.Name,
                            Int32.Parse(task.Duration.ToString()) / 480,
                            DateTime.Parse(dateStart),
                            DateTime.Parse(dateFinish),
                            task.ResourceNames, isSummary);
                        tasks.Add(Tasks);

                        tasksTable = Tasks.AddRow(tasksTable);
                    }
                    result = $"文件[{fileName}]读取成功！";
                }
                else
                {
                    result = $"文件[{fileName}]打开失败！";
                }
            }
            catch (System.Exception ex)
            {
                result = $"文件[{fileName}]读取失败.{newLine}{ex.Message}{newLine}{ex.StackTrace}";
            }
            finally
            {
                if (app != null)
                {
                    app.Quit(PjSaveType.pjDoNotSave);
                }
            }
            return result;
        }

        /// <summary>
        /// 编辑 MSProject 文件并保持
        /// </summary>
        /// <param name="taskName">任务名称</param>
        /// <param name="startDate">开始日期</param>
        /// <param name="finishDate">完成日期</param>
        /// <param name="taskId">任务Id</param>
        /// <returns>成功或失败，并自动保存编辑后的文件</returns>
        public string EditMSProject(string taskName, DateTime startDate, DateTime finishDate, int taskId)
        {
            ApplicationClass app = null;
            Project project;
            string result;

            try
            {
                app = new();
                if (app.FileOpen(fileName, readAndWrite, missing,
                    missing, missing, missing, missing,
                    missing, missing, missing, missing,
                    pool, missing, missing, missing, missing))
                {
                    project = app.ActiveProject;
                    app.Visible = false;
                    foreach (Task task in project.Tasks)
                    {
                        if (task.ID == taskId)
                        {
                            task.Name = taskName;
                            task.Start = startDate.ToString();
                            task.Finish = finishDate.ToString();
                            break;
                        }
                    }
                    project.SaveAs(fileName);
                    result = $"文件[{fileName}]编辑成功！";
                }
                else
                {
                    result = $"文件[{fileName}]打开失败！";
                }
            }
            catch (System.Exception ex)
            {
                result = $"文件[{fileName}]编辑失败.{newLine}{ex.Message}{newLine}{ex.StackTrace}";
            }
            finally
            {
                if (app != null)
                {
                    app.Quit(PjSaveType.pjDoNotSave);
                }
            }
            return result;
        }
    }
}
