/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  WordHelper
 * 描    述:
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 16:44:38
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 16:44:38
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

using CRCC.SystemTools.Log;
using Spire.Doc;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;

namespace CRCC.SystemTools.FileHelper
{
    public class WordHelper
    {
        [Description("更新指定模板文档中对应名称字段的值(注意：1，datas的Key为文档中已添加的MergeField域名称；Value为对应值，可以是Image对象。" +
           "2，注意制作正确格式的模板,如果是图片，可以在标签后预先插入一个空图片并设置图片格式。)")]
        public static Result<string> UpdateMergeFieldValues(string templatePath, Dictionary<string, object> datas, string savePath = null)
        {
            if (string.IsNullOrEmpty(templatePath) || File.Exists(templatePath) == false)
                return new Result<string>(false, null, "模板文件不存在！");
            try
            {
                var doc = new Document(templatePath);
                var fields = doc.Fields.OfType<MergeField>().ToList();
                foreach (var field in fields)
                {
                    var name = field.FieldName;
                    var text = field.Text;
                    if (text == name || text == "«" + name + "»")
                        field.Text = "";
                    if (string.IsNullOrEmpty(name) || datas.ContainsKey(name) == false)
                        continue;
                    var data = datas[name];
                    if (data is Image image)
                    {
                        var paragraph = field?.OwnerParagraph;
                        if (paragraph == null)
                            continue;
                        var index = paragraph.ChildObjects.IndexOf(field);
                        if (index < 0)
                            continue;
                        var picindex = index + 1;
                        var pic = picindex >= paragraph.ChildObjects.Count ? null : paragraph.ChildObjects[picindex] as DocPicture;
                        var w = pic == null ? image.Width : pic.Width;
                        var h = pic == null ? image.Height : pic.Height;
                        if (pic == null)
                        {
                            pic = new DocPicture(doc);
                            paragraph.ChildObjects.Insert(picindex, pic);
                        }
                        pic.LoadImage(image);
                        pic.Width = w;
                        pic.Height = h;
                    }
                    else
                        field.Text = data?.ToString();
                }
                doc.SaveToFile(savePath);
                return new Result<string>(true, savePath);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("更新文件中MergeField数据异常！", "更新文件中MergeField数据异常:" + ex);
                return new Result<string>(false, null, ex.Message);
            }
        }
    }
}