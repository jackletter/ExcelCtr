/* 版权所有：JACKOA 
 * 类 名 称：ExcelOP
 * 作    者：胡庆杰
 * 电子邮箱：1286317554@QQ.com
 * 创建日期：2016-03-04 
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data;
using System.Data.Common;
using System.Collections;

namespace ExcelCtr
{
    /// <summary>
    /// Excel操作类,用于控制excel的读取和写入
    /// </summary>
    public class ExcelOP
    {
        #region 读取excel
        /// <summary>将excel中的每一个表第一行为列名组合读取成一个dataset</summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static DataSet Read(string filePath)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath);
            return ds;
        }

        /// <summary>将excel中的每一个表第一行为列名组合读取成一个dataset</summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static DataSet Read(Stream stream)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream);
            return ds;
        }

        /// <summary>读取excel中指定表名和指定相应列头行的表</summary>
        /// <param name="strFileName"></param>
        /// <param name="sheetNames"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<string> sheetNames, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetNames, indexOfColNames);
            return ds;
        }

        /// <summary>读取excel中指定表名和指定相应列头行的表</summary>
        /// <param name="stream"></param>
        /// <param name="sheetNames"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<string> sheetNames, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetNames, indexOfColNames);
            return ds;
        }

        /// <summary>读取excel中指定表索引和相应列头行的表</summary>
        /// <param name="strFileName"></param>
        /// <param name="sheetIndexs"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<int> sheetIndexs, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetIndexs, indexOfColNames);
            return ds;
        }

        /// <summary>读取excel中指定表索引和相应列头行的表</summary>
        /// <param name="stream"></param>
        /// <param name="sheetIndexs"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<int> sheetIndexs, List<int> indexOfColNames)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetIndexs, indexOfColNames);
            return ds;
        }

        /// <summary>读取excel中指定表索引以及是否有列头行的读取数据情况</summary>
        /// <param name="filePath"></param>
        /// <param name="sheetIndexs"></param>
        /// <param name="hasColNames"></param>
        /// <param name="dataStartIndex"></param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<int> sheetIndexs, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetIndexs, hasColNames, dataStartIndex);
            return ds;
        }

        /// <summary>读取excel中指定表索引以及是否有列头行的读取数据情况</summary>
        /// <param name="stream"></param>
        /// <param name="sheetIndexs"></param>
        /// <param name="hasColNames"></param>
        /// <param name="dataStartIndex"></param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<int> sheetIndexs, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetIndexs, hasColNames, dataStartIndex);
            return ds;
        }

        /// <summary>读取excel中指定表名以及是否有列头行的读取数据情况</summary>
        /// <param name="filePath"></param>
        /// <param name="sheetNames"></param>
        /// <param name="hasColNames"></param>
        /// <param name="dataStartIndex"></param>
        /// <returns></returns>
        public static DataSet Read(string filePath, List<string> sheetNames, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(filePath, sheetNames, hasColNames, dataStartIndex);
            return ds;
        }

        /// <summary>读取excel中指定表名以及是否有列头行的读取数据情况</summary>
        /// <param name="stream"></param>
        /// <param name="sheetNames"></param>
        /// <param name="hasColNames"></param>
        /// <param name="dataStartIndex"></param>
        /// <returns></returns>
        public static DataSet Read(Stream stream, List<string> sheetNames, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = ExcelHelper.ImportExceltoDs(stream, sheetNames, hasColNames, dataStartIndex);
            return ds;
        }
        #endregion

        #region 写入excel
        /// <summary>将ds数据写入excel文件中</summary>
        /// <param name="filePath">生成excel文件的路径</param>
        /// <param name="ds">生成使用的数据集</param>
        public static void Write(string filePath, DataSet ds)
        {
            Write(filePath, ds, null);
        }

        /// <summary>将ds数据写入excel文件中</summary>
        /// <param name="filePath">生成excel文件的路径</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        public static void Write(string filePath, DataSet ds, List<string> SheetHeaders)
        {
            FileStream fs = new FileStream(filePath, FileMode.Create);
            Write(fs, ds, SheetHeaders);
        }

        /// <summary>将ds数据写入文件流中</summary>
        /// <param name="fs">目的文件流</param>
        /// <param name="ds">生成使用的数据集</param>
        public static void Write(FileStream fs, DataSet ds)
        {
            Write(fs, ds, null);
        }

        /// <summary>将ds数据写入文件流中</summary>
        /// <param name="fs">目的文件流</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        public static void Write(FileStream fs, DataSet ds, List<string> SheetHeaders)
        {
            Write(fs, ds, SheetHeaders, new List<string>());
        }

        /// <summary>将ds数据写入文件流中并指定合并行信息</summary>
        /// <param name="fs">目的文件流</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        /// <param name="combineColIndexs">要进行纵向合并的列索引集合</param>
        public static void Write(FileStream fs, DataSet ds, List<string> SheetHeaders, List<string> combineColIndexs)
        {
            MemoryStream stream = ExcelHelper.ExportDS(ds, SheetHeaders, combineColIndexs);
            byte[] bs = stream.ToArray();
            fs.Write(bs, 0, bs.Length);
            fs.Flush();
            fs.Close();
        }

        /// <summary>将ds数据写入excel文件中并指定合并行信息</summary>
        /// <param name="filePath">生成excel文件的路径</param>
        /// <param name="ds">生成使用的数据集</param>
        /// <param name="SheetHeaders">每个sheet的表头集合(顺序和ds的table对应)</param>
        /// <param name="combineColIndexs">要进行纵向合并的列索引集合</param>
        public static void Write(string filePath, DataSet ds, List<string> SheetHeaders, List<string> combineColIndexs)
        {
            FileStream fs = new FileStream(filePath, FileMode.Create);
            Write(fs, ds, SheetHeaders, combineColIndexs);
        }

        /// <summary>根据模板导出excel</summary>
        /// <param name="ht">传进去的参数</param>
        /// <param name="templatePath">模板配置文件的绝对路径,后缀名为.xml,注意仅支持97-2003格式Excel</param>
        public static void WriteWithTemplate(Hashtable ht, string templateConfPath, string destfilepath)
        {
            ExcelTemplateOP op = new ExcelTemplateOP(templateConfPath, ht);
            op.Write(destfilepath);
        }

        #endregion
    }
}