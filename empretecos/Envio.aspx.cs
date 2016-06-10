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

                    if (RadioButtonList_opcao.Items[1].Selected && TextBox_email.Text.IndexOf("@") < 1)
                    {
                        Label_result_upload.ForeColor = Color.Red;
                        Label_result_upload.Text = "Informe o e-mail.";
                        lOk = false;
                     
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
                        int n=-1;
                        int nLin = -1;
                        int col_name =0;
                        foreach (DataTable table in dsExcel.Tables)
                        {

                            //Localiza a coluna com o primeiro nome 10/06

                            foreach (DataColumn column_nome in table.Columns)
                            {
                                n++;
                                 if (table.Rows[0][column_nome].ToString().IndexOf("1.) Seu primeiro nome") >=0 && table.Columns.IndexOf(column_nome) > 2)
                                     col_name = n;
                                    {
                            } 
                                                      
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
                                               if (col_name > 0 ){
                                                pessoas[nLin, 0] = row[col_name].ToString();
                                               }else{
                                                pessoas[nLin, 0] = "";
                                               }
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
                                    if (RadioButtonList_opcao.Items[1].Selected)
                                    {
                                        MyMailMessage.To.Add(TextBox_email.Text);
                                    }
                                    else
                                    {
                                        MyMailMessage.To.Add(pessoas[nx, 3]);
                                        MyMailMessage.CC.Add("alexandre@prestus.com.br");

                                    }
                                    //MyMailMessage.To.Add(pessoas[nLin, 3]);
                                    MyMailMessage.Subject = "TECO: Pessoas interessadas na sua pergunta - Check List de Gestão PME";
                               
                                        cHtml = "";
                                        cHtml += "<html>";
                                        cHtml += "";
                                        cHtml += "<head>";
                                        cHtml += "<meta http-equiv=Content-Type content='text/html; charset=windows-1252'>";
                                        cHtml += "<meta name=Generator content='Microsoft Word 12 (filtered)'>";
                                        cHtml += "<style>";
                                        cHtml += "<!--";
                                        cHtml += "/* Font Definitions */";
                                        cHtml += "@font-face";
                                        cHtml += "{font-family:'Cambria Math'";
                                        cHtml += "panose-1:2 4 5 3 5 4 6 3 2 4;}";
                                        cHtml += "@font-face";
                                        cHtml += "{font-family:Calibri;";
                                        cHtml += "panose-1:2 15 5 2 2 2 4 3 2 4;}";
                                        cHtml += "/* Style Definitions */";
                                        cHtml += "p.MsoNormal, li.MsoNormal, div.MsoNormal";
                                        cHtml += "{margin-top:0cm;";
                                        cHtml += "margin-right:0cm;";
                                        cHtml += "margin-bottom:10.0pt;";
                                        cHtml += "margin-left:0cm;";
                                        cHtml += "line-height:115%;";
                                        cHtml += "font-size:11.0pt;";
                                        cHtml += "font-family:'Calibri','sans-serif';}";
                                        cHtml += "p";
                                        cHtml += "{margin-right:0cm;";
                                        cHtml += "margin-left:0cm;";
                                        cHtml += "font-size:12.0pt;";
                                        cHtml += "font-family:'Times New Roman','serif';}";
                                        cHtml += "span.apple-converted-space";
                                        cHtml += "{mso-style-name:apple-converted-space;}";
                                        cHtml += ".MsoPapDefault";
                                        cHtml += "{margin-bottom:10.0pt;";
                                        cHtml += "line-height:115%;}";
                                        cHtml += "@page WordSection1";
                                        cHtml += "{size:595.3pt 841.9pt;";
                                        cHtml += "margin:70.85pt 3.0cm 70.85pt 3.0cm;}";
                                        cHtml += "div.WordSection1";
                                        cHtml += "{page:WordSection1;}";
                                        cHtml += "-->";
                                        cHtml += "</style>";
                                        cHtml += "";
                                        cHtml += "</head>";
                                        cHtml += "";
                                        cHtml += "<body lang=PT-BR>";
                                        cHtml += "";
                                        cHtml += "<div class=WordSection1>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>Caro Empreteco,</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>&nbsp;</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>Acabamos de encerrar uma";
                                        cHtml += " fase do Check List de Gestão PME, onde as pessoas abaixo responderam SIM à";
                                        cHtml += " pergunta proposta por você.</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>Você deve agora fazer";
                                        cHtml += " contato individual com estes empretecos, e conhecer melhor suas necessidades.</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>&nbsp;</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>NÃO deixe de fazer";
                                        cHtml += " contato, pois de muitos que responderam, estes são os que declaram interesse na";
                                        cHtml += " melhoria (pergunta) que vc propôs!</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>Bons negócios!</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>&nbsp;</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><b><span style='color:#1F497D'>Grupo de Tecnologia e";
                                        cHtml += " Inovação (Coruja)</span></b></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>TECO</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>&nbsp;</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>O Check List de Gestão PME";
                                        cHtml += " é uma iniciativa realizada em crowdsourcing:</span></p>";
                                        cHtml += "<ul>";
                                        cHtml += "<li><p class=MsoNormal style='line-height:normal;background:white'><span";
                                        cHtml += "style='font-family:Symbol;color:#1F497D'></span><span style='font-size:7.0pt;";
                                        cHtml += "font-family:'Times New Roman','serif';color:#1F497D'></span><span";
                                        cHtml += "style='font-size:7.0pt;font-family:'Times New Roman','serif';color:#1F497D'>&nbsp;</span><span";
                                        cHtml += "style='color:#1F497D'>Idealização: Prestus (Alexandre Borin)</span></p>";
                                        cHtml += "<br/>";
                                        cHtml += "<li><p class=MsoNormal style='line-height:normal;background:white'><span";
                                        cHtml += "style='font-family:Symbol;color:#1F497D'></span><span style='font-size:7.0pt;";
                                        cHtml += "font-family:'Times New Roman','serif';color:#1F497D'></span><span";
                                        cHtml += "style='font-size:7.0pt;font-family:'Times New Roman','serif';color:#1F497D'>&nbsp;</span><span";
                                        cHtml += "style='color:#1F497D'>Desenvolvimento e Fluxo de informações: ERPSolutions";
                                        cHtml += " (Fabio Ponte / Grupo Coruja)</span></p>";
                                        cHtml += "<br/>";
                                        cHtml += "<li><p class=MsoNormal style='line-height:normal;background:white'><span";
                                        cHtml += "style='font-family:Symbol;color:#1F497D'></span><span style='font-size:7.0pt;";
                                        cHtml += "font-family:'Times New Roman','serif';color:#1F497D'></span><span";
                                        cHtml += "style='font-size:7.0pt;font-family:'Times New Roman','serif';color:#1F497D'>&nbsp;</span><span";
                                        cHtml += "style='color:#1F497D'>Colaboração: LuckInfo (Tania Souza) e Emilia Hiratuka";
                                        cHtml += " (Grupo Coruja)</span></p>";
                                        cHtml += "<br/>";
                                        cHtml += "<li><p class=MsoNormal style='line-height:normal;background:white'><span";
                                        cHtml += "style='font-family:Symbol;color:#1F497D'></span><span style='font-size:7.0pt;";
                                        cHtml += "font-family:'Times New Roman','serif';color:#1F497D'></span><span";
                                        cHtml += "style='font-size:7.0pt;font-family:'Times New Roman','serif';color:#1F497D'>&nbsp;</span><span";
                                        cHtml += "style='color:#1F497D'>Apoio: Comitê TECO</span></p>";
                                        cHtml += "<br/>";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>&nbsp;</span></p>";
                                        cHtml += "</ul>";
                                        cHtml += "<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:";
                                        cHtml += "normal;background:white'><span style='color:#1F497D'>=======</span></p>";
                                        cHtml += "";
                                        cHtml += "<p class=MsoNormal>&nbsp;</p>";
                                        cHtml += "";
                                        cHtml += "</div>";
                                        cHtml += "";
                                        cHtml += "</body>";
                                        cHtml += "";
                                    cHtml += "</html>";
                                    cHtml += "<table border='1'>";

                                  

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