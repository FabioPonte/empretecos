using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;
using System.Collections;
using System.IO;
using System.Drawing;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;



namespace empretecos
{
    public partial class Envio : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                    string path = System.Configuration.ConfigurationManager.AppSettings["upload"].ToString();


                    if (!System.IO.Directory.Exists(path)) ;
                    {
                        System.IO.Directory.CreateDirectory(path);

                    }
                    ArrayList listaArq = new ArrayList();
                    System.IO.DirectoryInfo dirCliente = new System.IO.DirectoryInfo(path);


                    foreach (System.IO.FileInfo arq in dirCliente.GetFiles())
                    {

                        listaArq.Add(arq.Name);

                    }


                    GridView_arquivos.DataSource = listaArq;

                    GridView_arquivos.DataBind();


           
            }
            catch (Exception ex)
            {
                Label_result_upload.ForeColor = Color.Red;
                Label_result_upload.Text = "Upload status: " + ex.Message;

            }
            
        }

        protected void Button_upload_Click(object sender, EventArgs e)
        {
            if (FileUploadControl.HasFile)
            {
                try
                {
                    Boolean lOk = true;
                    if (FileUploadControl.PostedFile.ContentLength > 3000000)
                    {
                        lOk = false;
                        Label_result_upload.Text = "Arquivo acima de 4MB";
                        Label_result_upload.ForeColor = Color.Red;
                        

                    }
                    if (lOk)
                    {
                        string path = System.Configuration.ConfigurationManager.AppSettings["upload"].ToString();
                        string path1 =path + @"\" + FileUploadControl.FileName;
                        path1 = path1.Replace(" ", "_");


                        System.IO.DirectoryInfo dirCliente = new System.IO.DirectoryInfo(path);
                        foreach (System.IO.FileInfo arq in dirCliente.GetFiles())
                        {
                            
                            if (arq.Name.ToLower().Contains(".xls") || arq.Name.ToLower().Contains(".xlsx") )
                            {
                                //File.Delete(arq.FullName);
                            }


                        }

                        FileUploadControl.SaveAs(path1.Replace("+", "_"));
                        ArrayList listaArq = new ArrayList();
                        System.IO.DirectoryInfo dirCliente2 = new System.IO.DirectoryInfo(path);


                        foreach (System.IO.FileInfo arq in dirCliente2.GetFiles())
                        {

                            listaArq.Add(arq.Name);

                        }


                        GridView_arquivos.DataSource = listaArq;

                        GridView_arquivos.DataBind();

                        Label_result_upload.ForeColor = Color.Blue;
                        Label_result_upload.Text = "Upload realizado!";
                        
                    }
                }
                catch (Exception ex)
                {
                    Label_result_upload.ForeColor = Color.Red;
                    Label_result_upload.Text = "Upload status: " + ex.Message;
           
                }
            }
            else
            {
                Label_result_upload.ForeColor = Color.Red;
                Label_result_upload.Text = "Selecione um arquivo";
              
            }
        }
    }
}