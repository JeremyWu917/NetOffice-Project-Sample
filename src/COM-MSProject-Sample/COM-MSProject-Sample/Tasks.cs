using System;
using System.Data;

namespace COM_MSProject_Sample
{
    /// <summary>
    /// MSProject 数据模型
    /// </summary>
    public class Tasks
    {
        /// <summary>
        /// 任务Id
        /// </summary>
        public int TaskId { get; set; }
        /// <summary>
        /// 任务名称
        /// </summary>
        public string ProjectName { get; set; }
        /// <summary>
        /// 工期
        /// </summary>
        public int DurationInDays { get; set; }
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime Start { get; set; }
        /// <summary>
        /// 完成时间
        /// </summary>
        public DateTime Finish { get; set; }
        /// <summary>
        /// 完成比例 0-1 标识(0% - 100%)
        /// </summary>
        public double PercentComplete { get; set; }
        /// <summary>
        /// 实际完成时间
        /// </summary>
        public DateTime ActualFinish { get; set; }
        /// <summary>
        /// 资源名称
        /// </summary>
        public string ResourceName { get; set; }
        /// <summary>
        /// 表单
        /// </summary>
        public DataTable Table { get; set; }
        /// <summary>
        /// 是否摘要/汇总
        /// </summary>
        public bool IsSummary { get; set; }
        /// <summary>
        /// 数据容器
        /// </summary>
        private DataSet DataSet;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="taskId"></param>
        /// <param name="projectName"></param>
        /// <param name="durationInDays"></param>
        /// <param name="start"></param>
        /// <param name="finish"></param>
        /// <param name="resourceName"></param>
        /// <param name="isSummary"></param>
        public Tasks(int taskId, string projectName, int durationInDays, DateTime start, DateTime finish, string resourceName, bool isSummary)
        {
            this.TaskId = taskId;
            this.ProjectName = projectName;
            this.DurationInDays = durationInDays;
            this.Start = start;
            this.Finish = finish;
            this.ResourceName = resourceName;
            this.IsSummary = isSummary;
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="table"></param>
        public Tasks(DataTable table)
        {
            this.Table = table;
        }

        /// <summary>
        /// 创建 DataTable 对象
        /// </summary>
        /// <returns></returns>
        public DataTable CreateTable()
        {
            Table = new DataTable("TasksTable");
            DataColumn column;

            column = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "ID",
                ReadOnly = true,
                Unique = true
            };
            Table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Task Name",
                ReadOnly = false,
                Unique = false
            };
            Table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "Duration",
                ReadOnly = false,
                Unique = false
            };
            Table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "Start",
                ReadOnly = false,
                Unique = false
            };
            Table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "Finish",
                ReadOnly = false,
                Unique = false
            };
            Table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Resource Names",
                ReadOnly = false,
                Unique = false
            };
            Table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "isSummary",
                ReadOnly = false,
                Unique = false
            };
            Table.Columns.Add(column);

            DataSet = new DataSet();
            DataSet.Tables.Add(Table);

            return Table;
        }

        /// <summary>
        /// 添加一列
        /// </summary>
        /// <param name="tableAddRows"></param>
        /// <returns></returns>
        public DataTable AddRow(DataTable tableAddRows)
        {
            DataRow row;

            row = tableAddRows.NewRow();
            row["ID"] = TaskId;
            row["Task Name"] = ProjectName;
            row["Duration"] = DurationInDays;
            row["Start"] = Start;
            row["Finish"] = Finish;
            row["Resource Names"] = ResourceName;
            row["isSummary"] = IsSummary;
            tableAddRows.Rows.Add(row);

            return tableAddRows;
        }

        /// <summary>
        /// 重写 ToString 方法
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return $"Project Name: {ProjectName} -- Duration: {DurationInDays} days -- Start Date: {Start} -- Finish Date: {Finish} -- Percent Complete: {PercentComplete} -- Resource Name: {ResourceName} -- Actual Finish Date: {ActualFinish} --\n\n";
        }
    }
}
