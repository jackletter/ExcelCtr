using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using DBUtil;
using System.Collections;
using System.Xml;
using System.Data;
using System.Reflection;

using NPOI.SS;
using NPOI.HSSF;
using NPOI;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;

using System.Text.RegularExpressions;
using ImageCode;

namespace ExcelCtr
{
    internal class ExcelTemplateOP
    {
        /// <summary>使用配置文件和哈希表(携带参数)初始化</summary>
        /// <param name="confPath">配置文件的绝对路径,以.xml结尾如:d:\demo.xml</param>
        /// <param name="ht">携带的参数</param>
        public ExcelTemplateOP(string templateConfPath, Hashtable ht)
        {
            this.templatePath = templateConfPath.Substring(0, templateConfPath.LastIndexOf('.')) + ".xls";
            ReadConf(templateConfPath);
            PrepareData(ht);
        }

        /// <summary>初始化配置文件</summary>
        /// <param name="ht">初始化配置携带的参数</param>
        private void PrepareData(Hashtable ht)
        {
            //先将声明好的参数传进来
            this.parameters.ForEach((i) =>
            {
                if (ht[i.name] != null)
                {
                    i.value = ht[i.name] ?? "";
                    ht.Remove(i.name);
                }
            });
            //将未声明的参数也传递进来
            ht.Keys.Cast<string>().ToList<string>().ForEach((i) =>
            {
                this.parameters.Add(new parameter()
                {
                    name = i,
                    receive = i,
                    type = (ht[i] ?? "").GetType().ToString(),
                    value = ht[i] ?? ""
                });
            });

            //初始化idb
            this.idbs.ForEach((i) =>
            {
                i.connstr_value = (i.connstr_conf ?? "").Trim(' ');
                if (i.connstr_value.StartsWith("parameters."))
                {
                    i.connstr_value = this.parameters
                        .FirstOrDefault<parameter>(ii => ii.name == i.connstr_value.Replace("parameters.", ""))
                        .value.ToString();
                }
                i.dbtype_value = (i.dbtype_conf ?? "").Trim(' ');
                if (i.dbtype_value.StartsWith("parameters."))
                {
                    i.dbtype_value = this.parameters
                        .FirstOrDefault<parameter>(ii => ii.name == i.dbtype_value.Replace("parameters.", ""))
                        .value.ToString();
                }
                i.value = IDBFactory.CreateIDB(i.connstr_value, i.dbtype_value);
            });

            //初始化计算结果表
            this.caldts.ForEach((i) =>
            {
                //先拿到iDb
                i.useidb_conf = i.useidb_conf ?? "";
                if (i.useidb_conf.StartsWith("parameters."))
                {
                    i.useidb_value = this.parameters
                        .FirstOrDefault<parameter>(ii => ii.name == i.useidb_conf.Replace("parameters.", ""))
                        .value as IDbAccess;
                }
                else if (i.useidb_conf.StartsWith("idbs."))
                {
                    i.useidb_value = this.idbs
                        .FirstOrDefault<idb>(ii => ii.name == i.useidb_conf.Replace("idbs.", ""))
                        .value as IDbAccess;
                }

                //获取para
                i.listpara.ForEach(ii =>
                {
                    ii.name = ii.name ?? "";
                    if (ii.name.StartsWith("parameters."))
                    {
                        parameter p = this.parameters.Single<parameter>(
                            iii => iii.name == ii.name.Replace("parameters.", ""));
                        ii.receive = p.receive;
                        ii.type = p.type;
                        ii.value = p.value;
                    }
                });

                //进行计算
                i.value = i.useidb_value
                    .GetDataTable(
                    string.Format(i.sqltmp,
                    i.listpara.Select<parameter, string>(ii => (ii.value ?? "").ToString()).ToArray()));

            });

            //初始化计算项
            this.calitems.ForEach((i) =>
            {
                if (string.IsNullOrWhiteSpace(i.from))
                {
                    //根据sql语句计算
                    #region
                    //1.先拿到iDb
                    i.useidb_conf = i.useidb_conf ?? "";
                    if (i.useidb_conf.StartsWith("parameters."))
                    {
                        parameter p = this.parameters
                            .FirstOrDefault<parameter>(ii => ii.name == i.useidb_conf.Replace("parameters.", ""));
                        if (p == null) throw new Exception("未找到数据库访问对象:" + i.useidb_conf);
                        i.useidb_value = p.value as IDbAccess;
                    }
                    else if (i.useidb_conf.StartsWith("idbs."))
                    {
                        idb p = this.idbs
                            .FirstOrDefault<idb>(ii => ii.name == i.useidb_conf.Replace("idbs.", ""));
                        if (p == null) throw new Exception("未找到数据库访问对象:" + i.useidb_conf);
                        i.useidb_value = p.value as IDbAccess;
                    }

                    //2.获取para
                    i.listpara.ForEach(ii =>
                    {
                        ii.name = ii.name ?? "";
                        if (ii.name.StartsWith("parameters."))
                        {
                            parameter p = this.parameters.Single<parameter>(
                                iii => iii.name == ii.name.Replace("parameters.", ""));
                            if (p == null) throw new Exception("未找到参数:" + ii.name);
                            ii.receive = p.receive;
                            ii.type = p.type;
                            ii.value = p.value;
                        }
                    });

                    //3.进行计算
                    i.value = i.useidb_value
                        .GetFirstColumnString(
                        string.Format(i.sqltmp,
                        i.listpara.Select<parameter, string>(ii => ii.value.ToString()).ToArray()));
                    #endregion
                }
                else
                {
                    //从计算表中引用的
                    #region
                    string from = i.from.Trim();
                    if (!from.StartsWith("caldts."))
                    {
                        throw new Exception(string.Format("计算项\"{0}\"的from属性\"{1}\"必须以\"caldts.\"开头", i.name, i.from));
                    }
                    string[] arr = from.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length < 3) throw new Exception(string.Format("计算项\"{0}\"的from属性\"{1}\"不符合规则,参照:\"caldts.JSYDYS.SJRQ\"", i.name, i.from));
                    caldt dt = this.caldts.FirstOrDefault<caldt>(j => j.name == arr[1]);
                    if (dt == null) throw new Exception(string.Format("计算项\"{0}\"的from属性\"{1}\"引用的计算表\"{2}\"未找到", i.name, i.from, arr[1]));
                    if (!dt.value.Columns.Contains(arr[2])) throw new Exception(string.Format("计算项\"{0}\"的from属性\"{1}\"引用的计算表\"{2}\"中未找到列\"{3}\"", i.name, i.from, arr[1], arr[2]));
                    DataRow[] rows = dt.value.Select();
                    if (!string.IsNullOrWhiteSpace(i.filter))
                    {
                        //根据filter筛选符合条件的行
                        rows = dt.value.Select(i.filter);
                    }
                    if (!string.IsNullOrWhiteSpace(i.fetch))
                    {
                        //根据fetch选取筛选后的行
                        if (!(i.fetch.Contains('[') &&
                            i.fetch.Contains(']') &&
                            i.fetch.Contains(':')))
                        {
                            throw new Exception(string.Format("计算项\"{0}\"的fetch属性\"{1}\"不符合规则,必须包含'[',']',':'三个字符,参照:\"[0:5]\",见python字符串截取语法", i.name, i.fetch));
                        }
                        List<DataRow> list = new List<DataRow>();
                        string[] fetcharr = i.fetch.Replace("[", "").Replace("]", "").Split(new char[] { ':' }, StringSplitOptions.None);
                        if (fetcharr.Length != 2)
                        {
                            throw new Exception(string.Format("计算项\"{0}\"的fetch属性\"{1}\"不符合规则,参照:\"[0:5]\",见python字符串截取语法", i.name, i.fetch));
                        }
                        if (string.IsNullOrWhiteSpace(fetcharr[0])) fetcharr[0] = "0";
                        if (string.IsNullOrWhiteSpace(fetcharr[1])) fetcharr[1] = rows.Length.ToString();
                        int start = int.Parse(fetcharr[0]);
                        int end = int.Parse(fetcharr[1]);
                        while (start < 0)
                        {
                            start += rows.Length;
                        }
                        while (end < 0)
                        {
                            end += rows.Length;
                        }
                        if (!(end < start || start > rows.Length - 1))
                        {
                            for (var k = start; k < rows.Length && k < end; k++)
                            {
                                list.Add(rows[k]);
                            }
                        }
                        rows = list.ToArray();
                    }
                    //根据聚合标志得到聚合后结果
                    if (string.IsNullOrWhiteSpace(i.aggregate))
                    {
                        i.aggregate = "str_join(,)";
                    }
                    if (i.aggregate.StartsWith("str_join"))
                    {
                        //字符串拼接
                        string joinstr = i.aggregate.Replace("str_join", "").Replace("(", "").Replace(")", "");
                        string _t = "";
                        for (var k = 0; k < rows.Length; k++)
                        {
                            string str = (rows[k][arr[2]] ?? "").ToString();
                            if (string.IsNullOrWhiteSpace(str)) continue;
                            if (k == 0)
                            {
                                _t += str;
                            }
                            else
                            {
                                _t += joinstr + str;
                            }
                        }
                        i.value = _t;
                    }
                    else if (i.aggregate == "sum")
                    {
                        //求和计算
                        double init = 0;
                        for (var k = 0; k < rows.Length; k++)
                        {
                            string str = (rows[k][arr[2]] ?? "").ToString();
                            double _t;
                            if (double.TryParse(str, out _t))
                            {
                                init += _t;
                            }
                        }
                        i.value = init.ToString();
                    }
                    else if (i.aggregate == "avg")
                    {
                        //不能转化为数字的不参与计算
                        double init = 0;
                        int len = 0;
                        for (var k = 0; k < rows.Length; k++)
                        {
                            string str = (rows[k][arr[2]] ?? "").ToString();
                            double _t;
                            if (double.TryParse(str, out _t))
                            {
                                len++;
                                init += _t;
                            }
                        }
                        i.value = (init / len).ToString();
                    }
                    else if (i.aggregate == "min")
                    {
                        //不能转化为数字的不参与计算
                        double init = 0;
                        for (var k = 0; k < rows.Length; k++)
                        {
                            string str = rows[k][arr[2]].ToString();
                            double _t;
                            if (double.TryParse(str, out _t))
                            {
                                init = init > _t ? _t : init;
                            }
                        }
                        i.value = init.ToString();
                    }
                    else if (i.aggregate == "max")
                    {
                        //不能转化为数字的不参与计算
                        double init = 0;
                        for (var k = 0; k < rows.Length; k++)
                        {
                            string str = rows[k][arr[2]].ToString();
                            double _t;
                            if (double.TryParse(str, out _t))
                            {
                                init = init > _t ? init : _t;
                            }
                        }
                        i.value = init.ToString();
                    }
                    else if (i.aggregate == "count")
                    {
                        //求数量
                        i.value = rows.Length.ToString();
                    }
                    #endregion
                }
            });
        }

        /// <summary>读取配置文件</summary>
        /// <param name="confPath"></param>
        private void ReadConf(string confPath)
        {
            string str = System.IO.File.ReadAllText(confPath, Encoding.UTF8);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(str);
            XmlElement root = doc.DocumentElement;
            if (root.Name != "WorkBook")
            {
                throw new Exception("配置文件的根节点必须是WorkBook");
            }
            if (string.IsNullOrEmpty(root.GetAttribute("version")))
            {
                throw new Exception("必须指定配置文件的版本");
            }
            XmlNodeList list = root.ChildNodes;
            IEnumerable<XmlElement> li = list.OfType<XmlElement>();
            li.ToList<XmlElement>().ForEach(i =>
            {
                int parameters_count = 0, idbs_count = 0, calitems_count = 0, caldts_count = 0, fastsheets_count = 0, sheets_count = 0;
                if (i.Name == "parameters" && parameters_count == 0)
                {
                    parameters_count++;
                    i.ChildNodes
                        .OfType<XmlElement>()
                        .Where<XmlElement>(ii => ii.Name == "parameter")
                        .ToList<XmlElement>()
                        .ForEach((iii) =>
                        {
                            this.parameters.Add(new parameter()
                            {
                                name = iii.GetAttribute("name"),
                                receive = iii.GetAttribute("receive"),
                                type = iii.GetAttribute("type")
                            });
                        });
                }
                if (i.Name == "idbs" && idbs_count == 0)
                {
                    idbs_count++;
                    i.ChildNodes.OfType<XmlElement>()
                        .Where<XmlElement>(ii => ii.Name == "idb")
                        .ToList<XmlElement>()
                        .ForEach((iii) =>
                        {
                            idb idb = new idb();
                            idb.name = iii.GetAttribute("name");
                            idb.connstr_conf = iii.ChildNodes
                                    .OfType<XmlElement>()
                                    .FirstOrDefault<XmlElement>(iiii => iiii.Name == "connstr")
                                    .GetAttribute("value");
                            idb.dbtype_conf = iii.ChildNodes
                                    .OfType<XmlElement>()
                                    .FirstOrDefault<XmlElement>(iiii => iiii.Name == "dbtype")
                                    .GetAttribute("value");
                            this.idbs.Add(idb);
                        });
                }
                if (i.Name == "calitems" && calitems_count == 0)
                {
                    calitems_count++;
                    i.ChildNodes.OfType<XmlElement>()
                        .Where<XmlElement>(ii => ii.Name == "calitem")
                        .ToList<XmlElement>()
                        .ForEach((iii) =>
                        {
                            calitem cal = new calitem();
                            this.calitems.Add(cal);
                            cal.name = iii.GetAttribute("name");
                            if (iii.HasAttribute("from"))
                            {
                                //该计算项是从计算表中引入的值
                                cal.from = iii.GetAttribute("from");
                                cal.fetch = iii.GetAttribute("fetch");
                                cal.filter = iii.GetAttribute("filter");
                                cal.aggregate = iii.GetAttribute("aggregate");
                            }
                            else
                            {
                                //该计算项是根据sql语句计算得来的
                                iii.ChildNodes.OfType<XmlElement>()
                                    .ToList<XmlElement>()
                                    .ForEach((iiii) =>
                                    {
                                        if (iiii.Name == "sqltmp" && string.IsNullOrEmpty(cal.sqltmp))
                                        {
                                            cal.sqltmp = iiii.InnerText.Trim(' ', '\t', '\r', '\n');
                                        }
                                        if (iiii.Name == "useidb" && string.IsNullOrEmpty(cal.useidb_conf))
                                        {
                                            cal.useidb_conf = iiii.GetAttribute("value");
                                        }
                                        if (iiii.Name == "usepara")
                                        {
                                            parameter p = this.parameters.FirstOrDefault<parameter>(para => para.name == iiii.GetAttribute("value").Replace("parameters.", ""));
                                            if (p == null) throw new Exception("未找到参数:" + iiii.GetAttribute("value"));
                                            cal.listpara.Add(p);
                                        }
                                    });
                            }
                        });
                }
                if (i.Name == "caldts" && caldts_count == 0)
                {
                    caldts_count++;
                    i.ChildNodes.OfType<XmlElement>()
                        .Where<XmlElement>(ii => ii.Name == "caldt")
                        .ToList<XmlElement>()
                        .ForEach((iii) =>
                        {
                            caldt cal = new caldt();
                            this.caldts.Add(cal);
                            cal.name = iii.GetAttribute("name");
                            iii.ChildNodes.OfType<XmlElement>()
                                .ToList<XmlElement>()
                                .ForEach((iiii) =>
                                {
                                    if (iiii.Name == "sqltmp" && string.IsNullOrEmpty(cal.sqltmp))
                                    {
                                        cal.sqltmp = iiii.InnerText.Trim(' ', '\t', '\r', '\n');
                                    }
                                    if (iiii.Name == "useidb" && string.IsNullOrEmpty(cal.useidb_conf))
                                    {
                                        cal.useidb_conf = iiii.GetAttribute("value");
                                    }
                                    if (iiii.Name == "usepara")
                                    {
                                        parameter p = this.parameters.FirstOrDefault<parameter>(para => para.name == iiii.GetAttribute("value").Replace("parameters.", ""));
                                        if (p == null) throw new Exception("未找到参数:" + iiii.GetAttribute("value"));
                                        cal.listpara.Add(p);
                                    }
                                });
                        });
                }
                if (i.Name == "fastsheets" && fastsheets_count == 0)
                {
                    fastsheets_count++;
                    this.fastsheets = i;
                }
                if (i.Name == "sheets" && sheets_count == 0)
                {
                    sheets_count++;
                    this.sheets = i;
                }
            });
        }

        /// <summary>将结果写入excel文件</summary>
        /// <param name="filePath"></param>
        public void Write(string destfilepath)
        {
            if (sheets == null && fastsheets == null) throw new Exception("模板文件【" + templatePath.Substring(0, templatePath.LastIndexOf('.')) + ".xml" + "】缺少sheets节点或fastsheets节点");

            if (fastsheets != null)
            {
                #region 优先解析fastsheets
                string useds = fastsheets.GetAttribute("useds");
                DataSet ds = null;
                List<string> SheetHeaders = new List<string>();
                List<string> combineColIndexs = new List<string>();

                #region 首先装载dataset
                if (!string.IsNullOrWhiteSpace(useds))
                {
                    //如果fastsheets使用了useds属性,就是用useds属性装载dataset
                    useds = useds.Trim(' ');
                    if (!useds.StartsWith("parameters.")) throw new Exception("模板文件【" + templatePath.Substring(0, templatePath.LastIndexOf('.')) + ".xml" + "】fastsheets节点的useds属性应该以\"parameters.\"开头");
                    ds = this.parameters.FirstOrDefault<parameter>(i => i.name == useds.Replace("parameters.", "")).value as DataSet;
                    if (ds == null) throw new Exception("未找到参数" + useds);
                }
                else
                {
                    //如果fastsheets没使用useds属性,就用fastsheet节点装载dataset
                    ds = new DataSet();
                    fastsheets.ChildNodes.OfType<XmlElement>()
                    .Where<XmlElement>(i => i.Name == "fastsheet")
                    .ToList<XmlElement>()
                    .ForEach(i =>
                    {
                        string usedt = i.GetAttribute("usedt");
                        usedt = (usedt ?? "").Trim(' ');
                        string name = i.GetAttribute("name");
                        name = (name ?? "").Trim(' ');
                        if (string.IsNullOrWhiteSpace(usedt)) throw new Exception("模板文件【" + templatePath.Substring(0, templatePath.LastIndexOf('.')) + ".xml" + "】fastsheets节点下的fastsheet节点缺少usedt属性");
                        DataTable dt = null;
                        if (usedt.StartsWith("parameters."))
                        {
                            usedt = usedt.Replace("parameters.", "");
                            dt = (this.parameters.FirstOrDefault<parameter>(ii => ii.name == usedt) ?? new parameter()).value as DataTable;
                            if (dt == null) throw new Exception("未找到表:parameters." + usedt);
                        }
                        else if (usedt.StartsWith("caldts."))
                        {
                            usedt = usedt.Replace("caldts.", "");
                            dt = (this.caldts.FirstOrDefault<caldt>(ii => ii.name == usedt) ?? new caldt()).value as DataTable;
                            if (dt == null) throw new Exception("未找到表:caldts." + usedt);
                        }

                        if (name != "")
                        {
                            dt.TableName = name;
                        }
                        ds.Tables.Add(dt);
                    });
                }
                #endregion

                #region 装载SheetHeaders和combineColIndexs
                fastsheets.ChildNodes.OfType<XmlElement>()
                    .Where<XmlElement>(i => i.Name == "fastsheet")
                    .ToList<XmlElement>()
                    .ForEach(i =>
                    {
                        string title = "";
                        string colindex = "";
                        i.ChildNodes.OfType<XmlElement>()
                            .ToList<XmlElement>()
                            .ForEach(ii =>
                            {
                                if (ii.Name == "title")
                                {
                                    if (!string.IsNullOrWhiteSpace(ii.GetAttribute("value")))
                                    {
                                        title = ii.GetAttribute("value");
                                    }
                                }
                                else if (ii.Name == "combineColIndexs")
                                {
                                    if (!string.IsNullOrWhiteSpace(ii.GetAttribute("value")))
                                    {
                                        string[] arrtmp = ii.GetAttribute("value").Trim(' ').Split(',');
                                        for (int j = 0; j < arrtmp.Length; j++)
                                        {
                                            colindex += GetColIndex(arrtmp[j]).ToString() + ",";
                                        }
                                        colindex = colindex.Trim(',');
                                    }
                                }
                            });
                        SheetHeaders.Add(title);
                        combineColIndexs.Add(colindex);
                    });
                #endregion

                //输出
                FileStream fs = new FileStream(destfilepath, FileMode.Create);
                MemoryStream stream = ExcelHelper.ExportDS(ds, SheetHeaders, combineColIndexs);
                byte[] bs = stream.ToArray();
                fs.Write(bs, 0, bs.Length);
                fs.Flush();
                fs.Close();
                #endregion
            }
            else if (sheets != null)
            {
                #region 不存在fastsheets节点的情况下解析sheets节点
                FileStream file = new FileStream(this.templatePath, FileMode.Open, FileAccess.Read);
                HSSFWorkbook book = new HSSFWorkbook(file);
                //解析sheets
                sheets.OfType<XmlElement>()
                    .Where<XmlElement>(i => i.Name == "sheet")
                    .ToList<XmlElement>()
                    .ForEach(sheet =>
                    {
                        ISheet isheet = book.GetSheet(sheet.GetAttribute("name"));
                        if (isheet == null) throw new Exception("模板excel【" + this.templatePath + "】中找不到sheet:" + sheet.GetAttribute("name"));
                        XmlElement rowmass = sheet.ChildNodes.OfType<XmlElement>()
                             .Where<XmlElement>(i => i.Name == "rowmass")
                             .FirstOrDefault<XmlElement>();
                        if (rowmass != null)
                        {
                            int currentrow = 0;
                            rowmass.ChildNodes.OfType<XmlElement>()
                                .Where<XmlElement>(i => i.Name == "row")
                                .ToList<XmlElement>()
                                .ForEach(row =>
                                {
                                    //拿到row节点下的model、position、index属性
                                    string model = row.GetAttribute("model");
                                    model = (model ?? "").Trim(' ');
                                    string position = row.GetAttribute("position");
                                    position = (position ?? "").Trim(' ');
                                    string index = row.GetAttribute("index");
                                    index = (index ?? "").Trim(' ');
                                    if (string.IsNullOrWhiteSpace(model)
                                        || string.IsNullOrWhiteSpace(position)
                                        || string.IsNullOrWhiteSpace(index)
                                        )
                                    {
                                        throw new Exception("标签row的属性model,position,index都不能为空");
                                    }
                                    //根据定位属性计算出当前应该操作的行索引
                                    if (position == "absolute")
                                    {
                                        currentrow = int.Parse(index) - 1;
                                    }
                                    else if (position == "relative")
                                    {
                                        currentrow += int.Parse(index);
                                    }

                                    if (model == "single")
                                    {
                                        #region 单行操作,不涉及到循环行
                                        row.ChildNodes.OfType<XmlElement>()
                                            .Where<XmlElement>(i => i.Name == "coltmp")
                                            .ToList<XmlElement>()
                                            .ForEach(col =>
                                            {
                                                string colindex = col.GetAttribute("index");
                                                string colcelltype = col.GetAttribute("celltype");
                                                colindex = (colindex ?? "").Trim(' ');
                                                if (string.IsNullOrWhiteSpace(colindex)) throw new Exception("coltmp标签的index属性不能为空!");
                                                string colval = col.GetAttribute("value") ?? "";
                                                string celltype = col.GetAttribute("type") ?? "";
                                                //colval = colval.Trim(' ');
                                                if (colval == "")
                                                {
                                                    colval = isheet.GetRow(currentrow).GetCell(GetColIndex(colindex), MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue ?? "";
                                                }
                                                //colval = colval.Trim(' ');
                                                if (colval != "")
                                                {
                                                    string res = ParseVal(colval);
                                                    if (celltype.ToLower() == "number")
                                                    {
                                                        double res_double;
                                                        if (double.TryParse(res, out res_double))
                                                        {
                                                            isheet.GetRow(currentrow).GetCell(GetColIndex(colindex), MissingCellPolicy.CREATE_NULL_AS_BLANK).SetCellValue(res_double);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        ICell cell = isheet.GetRow(currentrow).GetCell(GetColIndex(colindex), MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                                        ICellStyle cellStyle = cell.CellStyle;
                                                        if (cellStyle == null)
                                                        {
                                                            cellStyle = book.CreateCellStyle();
                                                            cell.CellStyle = cellStyle;
                                                        }
                                                        cellStyle.WrapText = true;
                                                        res = res.Replace("\\r", "\r")
                                                            .Replace("\\n", "\n")
                                                            .Replace("\\t", "\t");
                                                        cell.SetCellValue(new HSSFRichTextString(res));
                                                    }
                                                }

                                            });
                                        #endregion
                                    }
                                    else if (model == "cycle")
                                    {
                                        #region 循环行操作
                                        string binddt = row.GetAttribute("binddt");
                                        binddt = (binddt ?? "").Trim(' ');
                                        DataTable curdt = null;//存储当前行绑定到的DataTable
                                        #region 首先从caldts和parameters中解析出指定的DataTable
                                        if (binddt.StartsWith("caldts."))
                                        {
                                            string binddt_tmp = binddt.Replace("caldts.", "");
                                            caldt ctmp = this.caldts.Where<caldt>(i => i.name == binddt_tmp).FirstOrDefault<caldt>();
                                            if (ctmp == null) throw new Exception("循环行导出中未找到计算表项:" + binddt);
                                            curdt = ctmp.value;
                                        }
                                        else if (binddt.StartsWith("parameters."))
                                        {
                                            string binddt_tmp = binddt.Replace("parameters.", "");
                                            parameter para = this.parameters.Where<parameter>(i => i.name == binddt_tmp).FirstOrDefault<parameter>();
                                            if (para == null) throw new Exception("循环行导出中未找到参数项:" + binddt);
                                            if (para.value != null && para.value is DataTable)
                                            {
                                                curdt = para.value as DataTable;
                                            }
                                            else if (para.value is IList)
                                            {
                                                IList li = para.value as IList;
                                                Type type = para.value.GetType();
                                                Type[] tys = type.GetGenericArguments();
                                                if (tys.Length == 0) throw new Exception("从集合构建表过程中找不到类型参数,请检查要输出的集合数据！");
                                                Type inner = tys[0];
                                                PropertyInfo[] props = inner.GetProperties();
                                                DataTable dt = new DataTable();
                                                for (int i = 0; i < props.Length; i++)
                                                {
                                                    dt.Columns.Add(props[i].Name);
                                                }
                                                for (int j = 0; j < li.Count; j++)
                                                {
                                                    DataRow row_tmp = dt.NewRow();
                                                    for (int jj = 0; jj < props.Length; jj++)
                                                    {
                                                        row_tmp[dt.Columns[jj].ColumnName] = (props[jj].GetValue(li[j], null) ?? "").ToString();
                                                    }
                                                    dt.Rows.Add(row_tmp);
                                                }
                                                curdt = dt;
                                            }
                                            else
                                            {
                                                throw new Exception("无法根据参数加载参数项:" + binddt);
                                            }

                                        }
                                        #endregion

                                        if (curdt.Rows.Count == 0)
                                        {
                                            //如果绑定的DataTable中的记录数为0那就删除模板行
                                            if (isheet.LastRowNum >= currentrow + 1)
                                            {
                                                isheet.ShiftRows(currentrow + 1, isheet.LastRowNum, -1, true, false);
                                            }
                                            else
                                            {
                                                isheet.RemoveRow(isheet.GetRow(currentrow));
                                            }
                                            currentrow--;
                                        }
                                        else
                                        {

                                            List<string[]> coltmps = new List<string[]>();//存储模板列的配置参数,格式:0-索引,1-模板配置值,2-模板合并控制键

                                            #region 首先装载模板列的配置参数
                                            row.ChildNodes.OfType<XmlElement>()
                                                .Where<XmlElement>(i => i.Name == "coltmp")
                                                .ToList<XmlElement>()
                                                .ForEach(coltmp =>
                                                {
                                                    string coltmp_index = coltmp.GetAttribute("index") ?? "";
                                                    string coltmp_value = coltmp.GetAttribute("value") ?? "";
                                                    string coltmp_merge = coltmp.GetAttribute("mergekey") ?? "";
                                                    string coltmp_celltype = coltmp.GetAttribute("celltype") ?? "";
                                                    //模板列索引不能为空
                                                    if (coltmp_index == "")
                                                    {
                                                        throw new Exception("循环行的列模板coltmp标签的属性index不能为空.");
                                                    }
                                                    //模板列的引用,配置里找不到就去excel对应单元格中去找
                                                    if (coltmp_value == "")
                                                    {
                                                        coltmp_value = isheet.GetRow(currentrow).GetCell(GetColIndex(coltmp_index), MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue ?? "";
                                                    }
                                                    //模板引用不为空的话就添加存储
                                                    if (coltmp_value != "")
                                                    {
                                                        coltmps.Add(new string[] { coltmp_index, coltmp_value, coltmp_merge, coltmp_celltype });
                                                    }
                                                });
                                            #endregion

                                            int cyclestartrow_index = currentrow;//存储循环行的起始行索引
                                            //根据模板行和记录数插入缺少的行
                                            ExcelHelper.InsertRow(isheet, currentrow + 1, curdt.Rows.Count - 1, currentrow);

                                            for (int i = 0; i < curdt.Rows.Count; i++)
                                            {
                                                //解析当前行
                                                coltmps.ForEach(arr =>
                                                {
                                                    string[] res = ParseCycleVal(arr, curdt, i);
                                                    //输出列格式支持数字类型 2018-3-30
                                                    ICell cell = isheet.GetRow(currentrow).GetCell(GetColIndex(arr[0]), MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                                    if (arr[3] == "number")
                                                    {
                                                        double d_t;
                                                        if (double.TryParse(res[0], out d_t))
                                                        {
                                                            cell.SetCellValue(d_t);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        ICellStyle cellStyle = cell.CellStyle;
                                                        if (cellStyle == null)
                                                        {
                                                            cell.CellStyle = cellStyle;
                                                            cellStyle = book.CreateCellStyle();
                                                        }
                                                        cellStyle.WrapText = true;
                                                        res[0] = res[0].Replace("\\r", "\r")
                                                            .Replace("\\n", "\n")
                                                            .Replace("\\t", "\t");
                                                        cell.SetCellValue(new HSSFRichTextString(res[0]));
                                                    }
                                                    //如果存在控制合并键值,就进行预合并处理
                                                    if (arr[2] != "")
                                                    {
                                                        //将合并控制键对应的值填充进当前数据表中
                                                        if (!curdt.Columns.Contains(arr[2]))
                                                        {
                                                            curdt.Columns.Add(new DataColumn(arr[2]));
                                                        }
                                                        curdt.Rows[i][arr[2]] = res[1];
                                                    }
                                                });
                                                //当前行+1
                                                currentrow++;
                                            }
                                            //回到循环行的最后一行
                                            currentrow--;
                                            #region 纵向合并单元格
                                            coltmps.Where<string[]>(arr => arr[2] != "").ToList<string[]>()
                                            .ForEach(arr =>
                                            {
                                                if (curdt.Rows.Count > 1)//数据记录数大于1时才进行合并
                                                {
                                                    int curindex = cyclestartrow_index;//拿到循环行的起始行索引
                                                    //string val = (ExcelHelper.GetCellValue(isheet.GetRow(curindex).GetCell(GetColIndex(arr[0]))) ?? "").ToString();
                                                    string val = curdt.Rows[0][arr[2]].ToString();//拿到合并控制键对应的数据值
                                                    for (int i = 1; i < curdt.Rows.Count; i++)
                                                    {
                                                        //string realval = (ExcelHelper.GetCellValue(isheet.GetRow(cyclestartrow_index + i).GetCell(GetColIndex(arr[0]))) ?? "").ToString();
                                                        string realval = curdt.Rows[i][arr[2]].ToString();//拿到当前行合并控制键对应的数据值
                                                        if (realval == val)
                                                        {
                                                            //匹配成功
                                                            if (i == curdt.Rows.Count - 1)
                                                            {
                                                                //最后一行一定要参与合并
                                                                isheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(curindex, cyclestartrow_index + i, GetColIndex(arr[0]), GetColIndex(arr[0])));
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //匹配未成功
                                                            //如果之前处在匹配成功的状态里,那么进行合并操作
                                                            isheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(curindex, cyclestartrow_index + i - 1, GetColIndex(arr[0]), GetColIndex(arr[0])));
                                                            val = realval;
                                                            curindex = cyclestartrow_index + i;
                                                        }
                                                    }
                                                }
                                            });
                                            #endregion
                                        }
                                        #endregion
                                    }
                                });
                        }
                        #region 解析图片
                        sheet.OfType<XmlElement>()
                            .Where<XmlElement>(i => i.Name == "pic")
                            .ToList<XmlElement>()
                            .ForEach(pic =>
                            {
                                XmlElement from = pic.ChildNodes.OfType<XmlElement>()
                                     .Where<XmlElement>(i => i.Name == "from")
                                     .FirstOrDefault<XmlElement>();
                                if (from == null) throw new Exception("pic节点下找不到from节点");
                                string frommodel = (from.GetAttribute("model") ?? "").ToString().Trim(' ');
                                if (frommodel == "") throw new Exception("pic节点下的from节点的model属性不能为空");
                                byte[] bytes;
                                string fromvalue = (from.GetAttribute("value") ?? "").ToString().Trim(' ');
                                string res = "";
                                int index_c = 0;
                                Regex reg = new Regex(@"#(parameters|calitems)\.([^#]+)#");
                                Match mat = reg.Match(fromvalue);
                                string type = "";
                                string ext = "";
                                if (mat.Success)
                                {
                                    res += fromvalue.Substring(index_c, mat.Index - index_c);
                                    index_c = mat.Index + mat.Length;
                                    type = mat.Groups[1].Value;
                                    ext = mat.Groups[2].Value;
                                    if (type == "parameters")
                                    {
                                        parameter p = this.parameters.Where<parameter>(i => i.name == ext).FirstOrDefault<parameter>();
                                        if (p == null) throw new Exception("找不到参数:" + mat.Groups[0].Value);
                                        res += p.value.ToString();
                                    }
                                    else if (type == "calitems")
                                    {
                                        calitem cp = this.calitems.Where<calitem>(i => i.name == ext).FirstOrDefault<calitem>();
                                        if (cp == null) throw new Exception("找不到计算项:" + mat.Groups[0].Value);
                                        res += cp.value.ToString();
                                    }
                                }
                                while ((mat = mat.NextMatch()).Success)
                                {
                                    res += fromvalue.Substring(index_c, mat.Index - index_c);
                                    index_c = mat.Index + mat.Length;
                                    type = mat.Groups[1].Value;
                                    ext = mat.Groups[2].Value;
                                    if (type == "parameters")
                                    {
                                        parameter p = this.parameters.Where<parameter>(i => i.name == ext).FirstOrDefault<parameter>();
                                        if (p == null) throw new Exception("找不到参数:" + mat.Groups[0].Value);
                                        res += p.value.ToString();
                                    }
                                    else if (type == "calitems")
                                    {
                                        calitem cp = this.calitems.Where<calitem>(i => i.name == ext).FirstOrDefault<calitem>();
                                        if (cp == null) throw new Exception("找不到计算项:" + mat.Groups[0].Value);
                                        res += cp.value.ToString();
                                    }
                                }
                                res += fromvalue.Substring(index_c, fromvalue.Length - index_c);
                                if (frommodel == "QRCode")
                                {
                                    string qrsize = (from.GetAttribute("QRSize") ?? "").Trim(' ');
                                    int size = 100;
                                    if (qrsize != "") size = int.Parse(qrsize);
                                    //解析二维码                                    
                                    string filepath = Guid.NewGuid().ToString().Replace("-", "") + ".png";
                                    QRCodeOP.Encode(res, size, filepath, -1);
                                    bytes = File.ReadAllBytes(filepath);
                                    File.Delete(filepath);
                                    XmlElement stretch = pic.ChildNodes.OfType<XmlElement>()
                                        .Where<XmlElement>(i => i.Name == "stretch")
                                        .FirstOrDefault<XmlElement>();
                                    if (stretch == null) throw new Exception("pic节点下的必须存在stretch节点");
                                    XmlElement start = stretch.ChildNodes.OfType<XmlElement>().Where<XmlElement>(i => i.Name == "start")
                                        .FirstOrDefault<XmlElement>();
                                    if (start == null) throw new Exception("stretch节点下必须存在start节点");
                                    int col = GetColIndex(start.GetAttribute("col"));
                                    int row = int.Parse(start.GetAttribute("row")) - 1;
                                    int offx = int.Parse(start.GetAttribute("offx"));
                                    int offy = int.Parse(start.GetAttribute("offy"));
                                    //将图片数据装载到book中
                                    int picindex = book.AddPicture(bytes, PictureType.PNG);

                                    HSSFPatriarch patriarch = (HSSFPatriarch)isheet.CreateDrawingPatriarch();
                                    HSSFClientAnchor anchor = new HSSFClientAnchor(offx, offy, 0, 0, col, row, col, row);
                                    HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, picindex);
                                    pict.Resize();//设置图片按照原来的大小计算
                                }
                                else
                                {
                                    throw new Exception("仅支持二维码图片的插入,请将from节点的model配置项设置为QRCode");
                                }

                            });
                        #endregion
                    });
                //注意下面保存的写法,很重要,被坑了。。。
                FileStream sw = File.Create(destfilepath);
                book.Write(sw);
                sw.Close();
                #endregion
            }
        }

        /// <summary>解析值coltmp和pic\from的属性value的实际值
        /// </summary>
        /// <param name="colval">如:qwe#parameters.caseno#hjk</param>
        /// <returns></returns>
        private string ParseVal(string colval)
        {
            string res = "";
            int index_c = 0;
            Regex reg = new Regex(@"#(parameters|calitems)\.([^#]+)#");
            Match mat = reg.Match(colval);
            string type = "";
            string ext = "";
            if (mat.Success)
            {
                res += colval.Substring(index_c, mat.Index - index_c);
                index_c = mat.Index + mat.Length;
                type = mat.Groups[1].Value;
                ext = mat.Groups[2].Value;
                if (type == "parameters")
                {
                    parameter p = this.parameters.Where<parameter>(i => i.name == ext).FirstOrDefault<parameter>();
                    if (p == null) throw new Exception("找不到参数:" + mat.Groups[0].Value);
                    res += p.value.ToString();
                }
                else if (type == "calitems")
                {
                    calitem cp = this.calitems.Where<calitem>(i => i.name == ext).FirstOrDefault<calitem>();
                    if (cp == null) throw new Exception("找不到计算项:" + mat.Groups[0].Value);
                    res += cp.value.ToString();
                }
            }
            while ((mat = mat.NextMatch()).Success)
            {
                res += colval.Substring(index_c, mat.Index - index_c);
                index_c = mat.Index + mat.Length;
                type = mat.Groups[1].Value;
                ext = mat.Groups[2].Value;
                if (type == "parameters")
                {
                    parameter p = this.parameters.Where<parameter>(i => i.name == ext).FirstOrDefault<parameter>();
                    if (p == null) throw new Exception("找不到参数:" + mat.Groups[0].Value);
                    res += p.value.ToString();
                }
                else if (type == "calitems")
                {
                    calitem cp = this.calitems.Where<calitem>(i => i.name == ext).FirstOrDefault<calitem>();
                    if (cp == null) throw new Exception("找不到计算项:" + mat.Groups[0].Value);
                    res += cp.value.ToString();
                }
            }
            res += colval.Substring(index_c, colval.Length - index_c);
            return res;
        }

        /// <summary>解析循环行配置列的属性value的实际值,以及控制合并的值
        /// </summary>
        /// <param name="colval">如:qwe#parameters.caseno#hjk或#binddt.YueFen#月</param>
        /// <param name="curdt">循环行绑定的表</param>
        /// <param name="arr">模板列的配置数组</param>
        /// <param name="i">数据表curdt进行到的行索引</param>
        /// <returns></returns>
        private string[] ParseCycleVal(string[] arr, DataTable curdt, int i)
        {
            string[] res = new string[2];
            Regex reg = new Regex(@"#(parameters|calitems|binddt)\.([^#]+)#");
            for (int ii = 0; ii < 2; ii++)
            {
                if (arr[ii + 1] == "")
                {
                    res[ii] = "";
                    continue;
                }
                int index_c = 0;
                Match mat = reg.Match(arr[ii + 1]);
                string type = "";
                string ext = "";
                if (mat.Success)
                {
                    res[ii] += arr[ii + 1].Substring(index_c, mat.Index - index_c);
                    index_c = mat.Index + mat.Length;
                    type = mat.Groups[1].Value;
                    ext = mat.Groups[2].Value;
                    if (type == "parameters")
                    {
                        parameter p = this.parameters.Where<parameter>(pa => pa.name == ext).FirstOrDefault<parameter>();
                        if (p == null) throw new Exception("找不到参数:" + mat.Groups[0].Value);
                        res[ii] += p.value.ToString();
                    }
                    else if (type == "calitems")
                    {
                        calitem cp = this.calitems.Where<calitem>(ca => ca.name == ext).FirstOrDefault<calitem>();
                        if (cp == null) throw new Exception("找不到计算项:" + mat.Groups[0].Value);
                        res[ii] += cp.value.ToString();
                    }
                    else if (type == "binddt")
                    {
                        res[ii] += curdt.Rows[i][ext].ToString();
                    }
                }
                while ((mat = mat.NextMatch()).Success)
                {
                    res[ii] += arr[ii + 1].Substring(index_c, mat.Index - index_c);
                    index_c = mat.Index + mat.Length;
                    type = mat.Groups[1].Value;
                    ext = mat.Groups[2].Value;
                    if (type == "parameters")
                    {
                        parameter p = this.parameters.Where<parameter>(pa => pa.name == ext).FirstOrDefault<parameter>();
                        if (p == null) throw new Exception("找不到参数:" + mat.Groups[0].Value);
                        res[ii] += p.value.ToString();
                    }
                    else if (type == "calitems")
                    {
                        calitem cp = this.calitems.Where<calitem>(cal => cal.name == ext).FirstOrDefault<calitem>();
                        if (cp == null) throw new Exception("找不到计算项:" + mat.Groups[0].Value);
                        res[ii] += cp.value.ToString();
                    }
                    else if (type == "binddt")
                    {
                        res[ii] += curdt.Rows[i][ext].ToString();
                    }
                }
                res[ii] += arr[ii + 1].Substring(index_c, arr[ii + 1].Length - index_c);
            }
            return res;
        }

        /// <summary>存储列索引映射
        /// </summary>
        private static Hashtable ht_colmap = new Hashtable();

        /// <summary>静态代码块,初始化列索引映射
        /// </summary>
        static ExcelTemplateOP()
        {
            ht_colmap.Add('A', 1); ht_colmap.Add('a', 1);
            ht_colmap.Add('B', 2); ht_colmap.Add('b', 2);
            ht_colmap.Add('C', 3); ht_colmap.Add('c', 3);
            ht_colmap.Add('D', 4); ht_colmap.Add('d', 4);
            ht_colmap.Add('E', 5); ht_colmap.Add('e', 5);
            ht_colmap.Add('F', 6); ht_colmap.Add('f', 6);
            ht_colmap.Add('G', 7); ht_colmap.Add('g', 7);
            ht_colmap.Add('H', 8); ht_colmap.Add('h', 8);
            ht_colmap.Add('I', 9); ht_colmap.Add('i', 9);
            ht_colmap.Add('J', 10); ht_colmap.Add('j', 10);
            ht_colmap.Add('K', 11); ht_colmap.Add('k', 11);
            ht_colmap.Add('L', 12); ht_colmap.Add('l', 12);
            ht_colmap.Add('M', 13); ht_colmap.Add('m', 13);
            ht_colmap.Add('N', 14); ht_colmap.Add('n', 14);
            ht_colmap.Add('O', 15); ht_colmap.Add('o', 15);
            ht_colmap.Add('P', 16); ht_colmap.Add('p', 16);
            ht_colmap.Add('Q', 17); ht_colmap.Add('q', 17);
            ht_colmap.Add('R', 18); ht_colmap.Add('r', 18);
            ht_colmap.Add('S', 19); ht_colmap.Add('s', 19);
            ht_colmap.Add('T', 20); ht_colmap.Add('t', 20);
            ht_colmap.Add('U', 21); ht_colmap.Add('u', 21);
            ht_colmap.Add('V', 22); ht_colmap.Add('v', 22);
            ht_colmap.Add('W', 23); ht_colmap.Add('w', 23);
            ht_colmap.Add('X', 24); ht_colmap.Add('x', 24);
            ht_colmap.Add('Y', 25); ht_colmap.Add('y', 25);
            ht_colmap.Add('Z', 26); ht_colmap.Add('z', 26);
        }

        /// <summary>获取列的真正索引(0-based)
        /// </summary>
        /// <param name="colindex">配置中的索引如:A(返回0)或AB(返回26)</param>
        /// <returns></returns>
        private int GetColIndex(string colindex)
        {
            int res;
            if (int.TryParse(colindex, out res))
            {
                return res;
            }
            char[] arr = colindex.ToCharArray();
            for (int i = arr.Length - 1; i >= 0; i--)
            {
                res += (int)((int)ht_colmap[arr[i]] * Math.Pow(26, arr.Length - 1 - i));
            }
            return res - 1;
        }

        #region 属性
        public List<parameter> parameters = new List<parameter>();
        public List<idb> idbs = new List<idb>();
        public List<calitem> calitems = new List<calitem>();
        public List<caldt> caldts = new List<caldt>();

        public XmlElement sheets = null;//边解析边输出sheets节点
        public XmlElement fastsheets = null;//边解析边输出fastsheets节点
        public string templatePath = null;//保存模板的路径如:d:\demo.xls

        #endregion

        #region 模型类
        public class parameter
        {
            public string name { set; get; }
            public string receive { set; get; }
            public string type { set; get; }
            public object value { set; get; }
        }
        public class idb
        {
            public string name { set; get; }
            public string connstr_conf { set; get; }
            public string connstr_value { set; get; }
            public string dbtype_conf { set; get; }
            public string dbtype_value { set; get; }
            public IDbAccess value { set; get; }
        }

        public class calitem
        {
            public string name { set; get; }
            public string sqltmp { set; get; }
            public string useidb_conf { set; get; }
            public IDbAccess useidb_value = null;
            public List<parameter> listpara = new List<parameter>();
            public string value { set; get; }

            //引用计算表需要的属性
            public string from { set; get; }
            public string fetch { set; get; }
            public string filter { set; get; }
            public string aggregate { set; get; }
        }

        public class caldt
        {
            public string name { set; get; }
            public string sqltmp { set; get; }
            public string useidb_conf { set; get; }
            public IDbAccess useidb_value = null;
            public List<parameter> listpara = new List<parameter>();
            internal DataTable value { set; get; }
            public object this[int rowIndex, int colIndex]
            {
                get
                {
                    return value.Rows[rowIndex][colIndex];
                }
            }
            public object this[int rowIndex, string colName]
            {
                get
                {
                    return value.Rows[rowIndex][colName];
                }
            }
            public int rowCount
            {
                get { return value.Rows.Count; }
            }

            public int colCount
            {
                get { return value.Columns.Count; }
            }
        }
        #endregion
    }
}
