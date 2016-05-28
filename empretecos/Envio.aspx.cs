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
using System.Data;
using System.Data.OleDb;
using System.Net.Mail;




namespace empretecos
{
                    
    public partial class Envio : System.Web.UI.Page
    {
        string[] arquivos = new string[10000];
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                    string path = System.Configuration.ConfigurationManager.AppSettings["upload"].ToString();
                    int ny = -1;

                    if (!System.IO.Directory.Exists(path)) ;
                    {
                        System.IO.Directory.CreateDirectory(path);

                    }
                    ArrayList listaArq = new ArrayList();
                    System.IO.DirectoryInfo dirCliente = new System.IO.DirectoryInfo(path);


                    foreach (System.IO.FileInfo arq in dirCliente.GetFiles())
                    {
                        ny++;
                        arquivos[ny] = arq.Name;
                        listaArq.Add("Arquivo: "+arq.Name + "- Data: " + arq.CreationTime );

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
            string[,] pessoas=  new string[10000,4];;
                     
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

                    for (int nz = 0; nz < arquivos.Length && arquivos != null; nz++)
                    {
                        if (arquivos[nz] == FileUploadControl.FileName.Replace(" ", "_"))
                        {
                            Label_result_upload.ForeColor = Color.Red;
                            Label_result_upload.Text = "Esse arquivo já foi processado.";
                            lOk = false;
                        }
                        
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

                        DataSet ds_Data = new DataSet();
                        OleDbConnection oleCon = new OleDbConnection();

                        string strExcelFile = path1.Replace("+", "_");
                        oleCon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelFile + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;

                        string sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelFile + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
                        
                        string SpreadSheetName = "";

                        OleDbDataAdapter Adapter = new OleDbDataAdapter();
                        OleDbConnection conn = new OleDbConnection(sConnectionString);

                        string strQuery;
                        conn.Open();

                        int workSheetNumber = 0;

                        DataTable ExcelSheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                        SpreadSheetName = ExcelSheets.Rows[workSheetNumber]["TABLE_NAME"].ToString();

                        strQuery = "select * from [" + SpreadSheetName + "] ";
                        OleDbCommand cmd = new OleDbCommand(strQuery, conn);
                        Adapter.SelectCommand = cmd;
                        DataSet dsExcel = new DataSet();
                        Adapter.Fill(dsExcel);
                        // For each table in the DataSet, print the row values.
                        int n;
                        int nLin = -1;
                        foreach (DataTable table in dsExcel.Tables)
                        {
                           
                                foreach (DataColumn column in table.Columns)
                                {
                                   // processamento do cabeçalho 
                                    if (table.Rows[0][column].ToString().IndexOf("@") >=0 && table.Columns.IndexOf(column) > 2)
                                    {
                              //          pessoas = new string[10000,4];
                                        n = -1;

                                         foreach (DataRow row in table.Rows)
                                        {
                                            if (row[column].ToString().ToLower().IndexOf("sim") >= 0)
                                            {
                                                n++;
                                                nLin++;
                                                if (nLin == 0)
                                                {

                                                    Label_result_upload.Text += "<table border='1'>";
                                                    Label_result_upload.Text += "<tr><td colspan = '2' align='center'>Log de Processamento</td></tr>";
                                                    
                                                }

                                                //Log de Processamento
                                                if (n == 0)
                                                {
                                                    Label_result_upload.Text += "<tr><td colspan = '2'>Empreteco..: "+table.Rows[0][column].ToString() + "</td></tr>";
                                                    Label_result_upload.Text += "<tr><td colspan = '2'>Pessoas interessadas..: </td></tr>";
                                                    Label_result_upload.Text += " <tr><td>e-mail: </td><td> Telefone: </td></tr>";
                                                }
                           
                                                Label_result_upload.Text += " <tr><td>" + row[1].ToString() + "</td><td>" + row[2].ToString() + "</td></tr>";
                                                pessoas[nLin, 0] = row[0].ToString();
                                                pessoas[nLin, 1] = row[1].ToString();
                                                pessoas[nLin, 2] = row[2].ToString();
                                                pessoas[nLin, 3] = table.Rows[0][column].ToString();

                                            }
                                        }
                                    }
                                    
                                }
                           
                        }
                        if (nLin > 0)
                            Label_result_upload.Text += " </table> ";

                        conn.Close();
                        string cEmpreteco="";
                        string cHtml="x";
                        MailMessage email = new MailMessage();
                        SmtpClient client = new SmtpClient();
                        MailMessage MyMailMessage = new MailMessage();
                        SmtpClient SMTPServer = new SmtpClient("smtp.gmail.com");
                        SMTPServer.Port = 587;
                        SMTPServer.Credentials = new System.Net.NetworkCredential(System.Configuration.ConfigurationManager.AppSettings["mail"].ToString(), System.Configuration.ConfigurationManager.AppSettings["passmail"].ToString());
                        SMTPServer.EnableSsl = true;
                                                
                       for(int nx = 0; nx < pessoas.Length; nx++)
                        {
                            if (pessoas[nx, 3]==null)
                            {
                                break;
                            }
                            if (cEmpreteco != pessoas[nx, 3])
                            {
                                if (cHtml != "x")
                                {
                                    cHtml += "</table>";

                                    MyMailMessage.Body = cHtml;
                                    MyMailMessage.IsBodyHtml = true;

                                    SMTPServer.Send(MyMailMessage);
                                }
                                    email = new MailMessage();
                                    client = new SmtpClient();
                                    MyMailMessage = new MailMessage();
                                    //'From requires an instance of the MailAddress type
                                    MyMailMessage.From = new MailAddress(System.Configuration.ConfigurationManager.AppSettings["mail"].ToString());
                                    //'Create the SMTPClient object and specify the SMTP GMail server
                                    //MyMailMessage.To.Add("fabio@erpsolutions.com.br");
                                    MyMailMessage.To.Add(pessoas[nLin, 3]);
                                    MyMailMessage.Subject = "Empreteco - Pessoas interessadas na sua empresa";
                                    cHtml = "<table border='1'>";
                                    cHtml = cHtml + "<tr><td>e-mail</td><td>Telefone</td></tr>";
                                    cEmpreteco = pessoas[nx, 3];
                                      
                            }
                            cHtml = cHtml + "<tr><td>";
                            cHtml += pessoas[nx, 1];
                            cHtml = cHtml + "</td><td>";
                            cHtml += pessoas[nx, 2];
                            cHtml = cHtml + "</td></tr>";
                           
                            

                        }

                       if (cHtml != "x")
                       {
                           cHtml += "</table>";

                           MyMailMessage.Body = cHtml;
                           MyMailMessage.IsBodyHtml = true;

                           SMTPServer.Send(MyMailMessage);
                       }
                        foreach (System.IO.FileInfo arq in dirCliente2.GetFiles())
                        {

                            listaArq.Add("Arquivo: " + arq.Name + "- Data: " + arq.CreationTime );

                        }


                        GridView_arquivos.DataSource = listaArq;

                        GridView_arquivos.DataBind();

                        Label_result_upload.ForeColor = Color.Blue;
                        //Label_result_upload.Text = "Upload realizado!";
                        
                    }
                }
                catch (Exception ex)
                {
                    Label_result_upload.ForeColor = Color.Red;
                    Label_result_upload.Text = "Upload status: " +ex.Message+"-"+ ex.StackTrace;
           
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