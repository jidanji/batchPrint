using Microsoft.Office;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordPrint
{
    public class WriteIntoWord
    {
        ApplicationClass app = null;   //定义应用程序对象 
        Document doc = null;   //定义 word 文档对象
        Object missing = System.Reflection.Missing.Value; //定义空变量
        Object isReadOnly = false;
        // 向 word 文档写入数据 
        public void OpenDocument(string FilePath)
        {
            object filePath = FilePath;//文档路径
            app = new ApplicationClass(); //打开文档 
            doc = app.Documents.Open(ref filePath, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing);
            doc.Activate();//激活文档
        }
        /// <summary> 
        /// </summary> 
        ///<param name="parLableName">域标签</param> 
        /// <param name="parFillName">写入域中的内容</param> 
        /// 
        //打开word，将对应数据写入word里对应书签域

        public void WriteIntoDocument(string BookmarkName, string FillName)
        {
            if (BookmarkName == "Pic")
            {
                object obj2 = BookmarkName;
                doc.Bookmarks[ref obj2].Select();
                InlineShape shape1 = this.app.Selection.InlineShapes.AddPicture(FillName, ref missing, ref missing, ref missing);
                shape1.Width = 60f;
                shape1.Height = 80f;
                return;
            }

            object bookmarkName = BookmarkName;
            Bookmark bm = doc.Bookmarks.get_Item(ref bookmarkName);//返回书签 
            bm.Range.Text = FillName;//设置书签域的内容
        }
        /// <summary> 
        /// 保存并关闭 
        /// </summary> 
        /// <param name="parSaveDocPath">文档另存为的路径</param>
        /// 
        public void Save_CloseDocument(string SaveDocPath)
        {
            object savePath = SaveDocPath;  //文档另存为的路径 
            Object saveChanges = app.Options.BackgroundSave;//文档另存为 
            doc.SaveAs(ref savePath, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing);
            doc.Close(ref saveChanges, ref missing, ref missing);//关闭文档
            app.Quit(ref missing, ref missing, ref missing);//关闭应用程序

        }
    }
}
