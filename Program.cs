using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Reflection;
using System.Net.Mail;
using Newtonsoft.Json;
using ZohoDeskApi;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Zoho_TFS
{
    class Program
    {
        const string ZOHODESK_API_URL = "https://desk.zoho.com/api/v1";
        const string ZOHODESK_AUTH_TOKEN = "";
        const string ZOHODESK_ORG_ID = "";

        const string TFS_URL = "http://spurs:8080/tfs/";
        const string TFS_PROJECT = "VNext";

        const string MAIL_HOST = "smtp.gmail.com";
        const string MAIL_USER = "";
        const string MAIL_PASS = "";

        const int MINUTES = 10;

        static void Main(string[] args)
        {
            while (true)
            {
                deleteAllTempFiles();

                copyNewTicketsToTFS();

                setTagsToTFSTasks();

                setTFSVersionToZohoTickets();

                sendTicketsToTRSistemas();
                
                //getAllDevelopedVersionedTickets();

                Console.WriteLine("* Ultimo Ciclo: " + DateTime.Now + " *");

                // Tempo em MINUTOS para entrar na próxima iteração
                System.Threading.Thread.Sleep(MINUTES * 60000);
            }
        }

        private static void sendTicketsToTRSistemas()
        {
            try
            {
                // Retorna os últimos tíckets que estão com o status "Encaminhado"
                Console.WriteLine("### Verificando lista de TR Enviado ### : " + DateTime.Now);

                var lastTickets = JsonConvert.DeserializeObject<ZohoDeskApi.TicketList>(
                    ZohoDeskGetData(ZOHODESK_API_URL + "/tickets/search?status=Aberto&limit=100&customField1=Módulo:TR*")).data.ToList();

                // Lê cada um dos tickets encaminhados
                foreach (var ticketItem in lastTickets)
                {
                    Console.WriteLine("\nLendo ticket #" + ticketItem.ticketNumber);

                    // Carregar os detalhes do ticket
                    var currentTicket = ticketItem;

                    Console.WriteLine(currentTicket.classification);

                    string _subj = "[" + currentTicket.classification + "] #" + currentTicket.ticketNumber + " " + currentTicket.subject;
                    string _body =
                        @"[DESCRICAO]<br>" + currentTicket.description + "<br><br>" +
                        @"[AGENTE DE SUPORTE]<br>" + NullToString(currentTicket.assignee.firstName) + " " + NullToString(currentTicket.assignee.lastName) + "<br><br>" +
                        @"[CONTATO DO CLIENTE]<br>" + NullToString(currentTicket.contact.firstName) + " " + NullToString(currentTicket.contact.lastName) + "<br><br>" +
                        @"[PRIORIDADE DO CLIENTE]<br>" + currentTicket.priority;

                    sendEmail(_subj, _body);

                    updateTicketStatus(currentTicket.id, "TR Enviado");

                }

            } catch { return; }

        }

        private static void sendEmail(string subject, string body)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();

                SmtpClient = new SmtpClient();
                SmtpClient.Host = MAIL_HOST;
                SmtpClient.Port = 587;
                SmtpClient.EnableSsl = true;
                SmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                SmtpClient.UseDefaultCredentials = false;
                SmtpClient.Credentials = new NetworkCredential(MAIL_USER, MAIL_PASS);

                MailMessage MailMessage = new MailMessage();
                
                MailMessage = new MailMessage();
                MailMessage.From = new MailAddress("teste@gmail.com", "Teste", Encoding.UTF8);
                // Envia e-mail para a TR SISTEMAS
                MailMessage.To.Add(new MailAddress("outroteste@gmail.com", "Outro Teste", Encoding.UTF8));
                // Envia cópia para membros da EDAX

                MailMessage.Subject = subject;
                MailMessage.IsBodyHtml = true;
                MailMessage.Body = body;
                MailMessage.BodyEncoding = Encoding.UTF8;
                MailMessage.BodyEncoding = Encoding.GetEncoding("ISO-8859-1");

                MailMessage.Priority = MailPriority.High;

                SmtpClient.Send(MailMessage);
            }
            catch (SmtpFailedRecipientException ex)
            {
                Console.WriteLine("Mensagem : {0} " + ex.Message);
            }
            catch (SmtpException ex)
            {
                Console.WriteLine("Mensagem SMPT Fail : {0} " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Mensagem Exception : {0} " + ex.Message);
            }
        }

        private static void getAllDevelopedVersionedTickets()
        {
            Console.WriteLine("### Verificando lista de Encaminhados ### : " + DateTime.Now);

            var lastTickets = JsonConvert.DeserializeObject<ZohoDeskApi.TicketList>(
                ZohoDeskGetData(ZOHODESK_API_URL + "/tickets/search?status=Fechado&limit=100&customField1=ReleasedVersion:v*")).data.ToList();

            using (StreamWriter writer = File.CreateText("newfile.csv"))
            {
                foreach (Ticket item in lastTickets)
                {
                    writer.WriteLine(
                        item.contact.email + ";" +
                        item.classification + ";" +
                        "#" + item.ticketNumber + ";" +
                        item.subject + ";" +
                        item.customFields.releasedVersion + ";" +
                        item.status + ";" +
                        item.customFields.módulo + ";" +
                        item.priority);
                }
            }
        }

        /// <summary>
        /// Varre diretório onde estão os arquivos de versão no formato
        /// 0.0.0.txt e busca todos os IDs das tasks contidos nele
        /// Adiciona a tag da versão correspondente a cada uma das tasks
        /// </summary>
        private static void setTagsToTFSTasks()
        {
            // Procura arquivo de versão na pasta /version
            string localPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\version\";
            DirectoryInfo di = new DirectoryInfo(localPath);

            foreach (FileInfo file in di.GetFiles())
            {
                // O nome do arquivo determina qual será a versão da TAG, ex: 0.0.0.txt
                var version = "v[" + file.Name.Replace(".txt", "") + "]";

                // Abre o arquivo e lê em cada linha qual o id do workitem correspondente no TFS
                string workItem;
                StreamReader fileStr = new StreamReader(file.FullName);

                while ((workItem = fileStr.ReadLine()) != null)
                {
                    Console.WriteLine(workItem);
                    // Adiciona a TAG da versão no work item no padrão v[0.0.0]
                    setTFSTag(workItem, version);
                }

                fileStr.Close();

                // Apaga o arquivo
                file.Delete();
            }


        }

        /// <summary>
        /// Adiciona uma nova TAG a uma task existente no TFS
        /// </summary>
        /// <param name="id">id da task no TFS</param>
        /// <param name="version"></param>
        private static void setTFSTag(string id, string version)
        {
            // Abrindo conexão com o TFS
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(TFS_URL));
            WorkItemStore wis = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));

            // Buscando todos os tickets do projeto VNext do tipo BUG
            WorkItemCollection wic = wis.Query("SELECT * FROM WorkItems WHERE [System.Id] = '" + id + "'");

            // Verificando se o ticket existe
            foreach (WorkItem item in wic)
            {
                // O ticket existe: Verificar status no TFS
                Console.WriteLine(" ### ENCONTRADO ### ");
                Console.WriteLine(item.Id + ": " + item.Title + " - [" + item.State + "]");
                
                foreach (var irl in item.Links)
                {
                    if (irl.GetType().Name.Equals("RelatedLink"))
                    {
                        RelatedLink _irl = (RelatedLink)irl;

                        if (_irl.LinkTypeEnd.Name.Equals("Parent"))
                        {
                            Console.WriteLine(_irl.RelatedWorkItemId);
                            setTFSTag(_irl.RelatedWorkItemId.ToString(), version);
                        }
                    }
                }

                item.Open();
                item.Fields["Tags"].Value += ";" + version;
                item.Save();
            }
        }

        /// <summary>
        /// Pega o padrão de versão na TAG do TFS v[0.0.0] 
        /// e adiciona um comentário no ticket correspondente do Zoho Desk 
        /// indicando ao cliente em que versão a demanda estará disponível
        /// </summary>
        private static void setTFSVersionToZohoTickets()
        {
            // Lista todos os tickets que estão na coluna de desenvolvidos
            Console.WriteLine("### Verificando lista de Desenvolvidos ### : " + DateTime.Now);

            var lastTickets = JsonConvert.DeserializeObject<ZohoDeskApi.TicketList>(
                ZohoDeskGetData(ZOHODESK_API_URL + "/tickets/search?status=Desenvolvido&limit=100")).data.ToList();

            foreach (var ticketItem in lastTickets)
            {
                Console.WriteLine("\nLendo ticket #" + ticketItem.ticketNumber);

                // Procura cada um dos tickets no TFS
                WorkItem tfsItem = getTicketOnTFS(ticketItem);
                if (tfsItem != null)
                {
                    // Verifica se o status da task no tpf é Done
                    if (tfsItem.State.Equals("Done"))
                    {
                        // Varre cada uma das TAGS
                        foreach (var tagItem in tfsItem.Tags.Split(';').ToList())
                        {
                            // Verifica a TAG que se refere a versão
                            if (!tagItem.Equals("") && tagItem.Trim().Substring(0, 2).Equals("v["))
                            {
                                var rv = (ticketItem.customFields.releasedVersion == null) ? "" : ticketItem.customFields.releasedVersion.Trim();
                                // Adiciona comentário no Ticket do Zoho Desk informando a versão em que a resolução do ticket estará disponível
                                // apenas se a versão informada anteriormente for maior do que a já conhecida
                                if (tagItem.Trim().CompareTo(rv) > 0)
                                    sendVersionReply(ticketItem, tagItem.Trim());
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Adiciona do Desk um comentário com a versão atualizada no TFS para o cliente do ticket
        /// </summary>
        /// <param name="ticketItem"></param>
        /// <param name="version"></param>
        private static void sendVersionReply(Ticket ticketItem, string version)
        {
            // https://desk.zoho.com/api/v1/tickets/84725000002829007/comments
            var request = (HttpWebRequest)WebRequest.Create(ZOHODESK_API_URL + "/tickets/" + ticketItem.id + "/comments");

            var postData = "{ " +
                "\"isPublic\" : \"true\", " +
                "\"content\" : \"A atualizacao disponivel a partir da versao: " + version + "\" " +
                "}";

            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            updateTicketVersionField(ticketItem, version);
        }

        /// <summary>
        /// Adiciona do Desk um comentário com progresso atualizado no TFS para o cliente do ticket
        /// </summary>
        /// <param name="ticketItem"></param>
        /// <param name="progress"></param>
        /// <param name="total"></param>
        private static void sendProgressReply(Ticket ticketItem, int progress, int total)
        {
            // https://desk.zoho.com/api/v1/tickets/84725000002829007/comments
            var request = (HttpWebRequest)WebRequest.Create(ZOHODESK_API_URL + "/tickets/" + ticketItem.id + "/comments");

            var postData = "{ " +
                "\"isPublic\" : \"true\", " +
                "\"content\" : \"Sua demanda esta em desenvolvimento. Ja concluimos " + progress + " etapa(s) de um total de " + total + ".\"" +
                "}";

            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            updateTicketProgressField(ticketItem, progress, total);
        }

        /// <summary>
        /// Atualiza o progresso no campo customizado do ticket do Zoho Desk
        /// </summary>
        /// <param name="ticketItem"></param>
        /// <param name="progress"></param>
        /// <param name="total"></param>
        private static void updateTicketProgressField(Ticket ticketItem, int progress, int total)
        {
            var request = (HttpWebRequest)WebRequest.Create(ZOHODESK_API_URL + "/tickets/" + ticketItem.id);

            var postData = "{ \"customFields\": { \"Progress\": \"" + progress + "/" + total + "\" } }";

            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "PATCH";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
        }

        /// <summary>
        ///   Atualiza o campo referente a versão ReleasedVersion no Zoho Desk
        /// </summary>
        /// <param name="ticketItem"></param>
        /// <param name="version"></param>
        private static void updateTicketVersionField(Ticket ticketItem, string version)
        {
            var request = (HttpWebRequest)WebRequest.Create(ZOHODESK_API_URL + "/tickets/" + ticketItem.id);

            var postData = "{ \"customFields\": { \"ReleasedVersion\": \"" + version + "\" } }";

            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "PATCH";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
        }

        /// <summary>
        /// Apaga os arquivos temporários
        /// </summary>
        private static void deleteAllTempFiles()
        {
            string localPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\temp\";
            DirectoryInfo di = new DirectoryInfo(localPath);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        /// <summary>
        /// Modifica o status de um ticket no Zoho Desk
        /// </summary>
        /// <param name="id">id do ticket no desk</param>
        /// <param name="status">texto que descreve o status do desk</param>
        private static void updateTicketStatus(string id, string status)
        {
            var request = (HttpWebRequest)WebRequest.Create(ZOHODESK_API_URL + "/tickets/" + id);

            var postData = "{ \"status\" = \"" + status + "\" }";
            //postData += "&thing2=world";

            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "PATCH";
            request.ContentType = "application/json";
            request.ContentLength = data.Length;
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

        }

        /// <summary>
        /// Realiza uma busca nos tickets do Zoho Desk,
        /// caso existam tickets no Desk que não existem no TFS
        /// é criada uma nova task no TFS.
        /// </summary>
        private static void copyNewTicketsToTFS()
        {
            // Retorna os últimos tíckets que estão com o status "Encaminhado"
            Console.WriteLine("### Verificando lista de Encaminhados ### : " + DateTime.Now);

            var lastTickets = JsonConvert.DeserializeObject<ZohoDeskApi.TicketList>(
                ZohoDeskGetData(ZOHODESK_API_URL + "/tickets/search?status=Encaminhado&limit=100")).data.ToList();

            // Lê cada um dos tickets encaminhados
            foreach (var ticketItem in lastTickets)
            {
                Console.WriteLine("\nLendo ticket #" + ticketItem.ticketNumber);

                // Carregar os detalhes do ticket
                var currentTicket = ticketItem;

                Console.WriteLine(currentTicket.classification);
                // Verificar se a classificação do ticket é Erro, Erro em Homologação ou Nova funcionalidade
                if (currentTicket.classification.Equals("Erro em homologação") ||
                    currentTicket.classification.Equals("Erro") ||
                    currentTicket.classification.Equals("Nova funcionalidade"))
                {
                    // Carrega as Threads do Ticket
                    var threadList = JsonConvert.DeserializeObject<ZohoDeskApi.ThreadList>(
                        ZohoDeskGetData(ZOHODESK_API_URL + "/tickets/" + ticketItem.id + "/threads")).data.ToList();

                    currentTicket.threadList = new ThreadList();

                    foreach (var threadItem in threadList)
                    {
                        // Verifica se a thread tem algum arquivo anexo
                        if (threadItem.hasAttach.Equals("true"))
                        {
                            // Carrega informações dos arquivos anexados nas threads
                            var attachmentList = JsonConvert.DeserializeObject<ZohoDeskApi.Thread>(
                                ZohoDeskGetData(ZOHODESK_API_URL + "/tickets/" + ticketItem.id + "/threads/" + threadItem.id)).attachments.ToList();

                            threadItem.attachments = attachmentList;
                        }

                        currentTicket.threadList.data = threadList;
                    }

                    // Verificar se o ticket já existe no TFS
                    WorkItem tfsItem = getTicketOnTFS(currentTicket);
                    if (tfsItem == null)
                    {
                        // O ticket não existe: Criar o ticket no TFS
                        setNewTicketOnTFS(currentTicket);
                    }
                    else
                    {
                        // O ticket existe: verificar se está com state Done
                        if (tfsItem.State.Equals("Done"))
                        {
                            // Muda o status do ticket no Desk
                            updateTicketStatus(ticketItem.id, "Desenvolvido");
                        }
                        else
                        {
                            // Verifica o andamento das subatividades dos backlog item
                            if (tfsItem.Type.Name.Equals("Product Backlog Item"))
                            {
                                updateTicketProgress(tfsItem, ticketItem);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Atualiza o progresso do ticket no Zoho Desk de acordo com as informações do TFS
        /// </summary>
        /// <param name="tfsWorkItem">Item do TFS</param>
        /// <param name="ticketId">ID do ticket no Zoho Desk</param>
        private static void updateTicketProgress(WorkItem tfsWorkItem, Ticket ticket)
        {
            if (tfsWorkItem.Links.Count > 0)
            {
                int progress = 0;
                int total = 0;
                // Verifica cada um dos workitens filho
                for (int i = 0; i < tfsWorkItem.Links.Count; i++)
                {
                    if (tfsWorkItem.Links[i].GetType().Name.Equals("RelatedLink"))
                    {
                        total++;
                        // Busca o workitem no TFS pelo Id se estiver concluido aumenta o progresso
                        if (getTicketOnTFS(((RelatedLink)tfsWorkItem.Links[i]).RelatedWorkItemId).State.Equals("Done") ||
                            getTicketOnTFS(((RelatedLink)tfsWorkItem.Links[i]).RelatedWorkItemId).State.Equals("Closed"))
                        { progress += 1; }
                    }
                }

                // Atualiza a informação do progresso no ticket do zoho
                string prg = progress + "/" + total;
                
                if (NullToString(ticket.customFields.progress).Equals("") || !ticket.customFields.progress.Equals(prg))
                {
                    sendProgressReply(ticket, progress, total);

                    Console.WriteLine("Andamento: {0} atividade(s) de {1} concluida(s)", progress, total);

                    // Se todas as atividades foram concluídas, muda o state do backlog item para done
                    if (progress == total)
                    {
                        tfsWorkItem.Open();
                        tfsWorkItem.State = "Done";
                        tfsWorkItem.Save();
                    }
                }
            }
        }

        /// <summary>
        /// Busca um workitem no TFS pelo Id
        /// </summary>
        /// <param name="relatedWorkItemId"></param>
        /// <returns></returns>
        private static WorkItem getTicketOnTFS(int relatedWorkItemId)
        {
            // Abrindo conexão com o TFS
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(TFS_URL));
            WorkItemStore wis = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));

            // Buscando todos os tickets do projeto VNext do tipo BUG
            WorkItemCollection wic = wis.Query("SELECT * FROM WorkItems WHERE [System.Id] = '" + relatedWorkItemId + "'");

            // Verificando se o ticket existe
            foreach (WorkItem item in wic)
            {
                // O ticket existe: Verificar status no TFS
                Console.WriteLine(" ### ENCONTRADO ### ");
                Console.WriteLine(item.Id + ": " + item.Title + " - [" + item.State + "]");

                return item;
            }

            // Não existe Bug
            return null;
        }

        /// <summary>
        /// Copia um ticket do Zoho no TFS
        /// </summary>
        /// <param name="currentTicket">Ticket</param>
        private static void setNewTicketOnTFS(Ticket currentTicket)
        {
            var wiType = "Product Backlog Item";
            var wiField = "Description";
            if (currentTicket.classification.Equals("Erro") || 
                currentTicket.classification.Equals("Erro em homologação"))
            {
                wiType = "Bug";
                wiField = "Repro Steps";
            }

            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(TFS_URL));
            WorkItemStore wis = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));

            Project teamProject = wis.Projects[TFS_PROJECT];
            WorkItemType witBug = teamProject.WorkItemTypes[wiType];

            if (currentTicket.assignee == null)
            {
                currentTicket.assignee = new Assignee();
                currentTicket.assignee.firstName = "Time";
                currentTicket.assignee.lastName = "de Suporte";
            }

            if (witBug != null)
            {
                Console.WriteLine("Copiando o ticket #{0} para o TFS!", currentTicket.ticketNumber);

                WorkItem wi = new WorkItem(witBug);

                wi.Title = "#" + currentTicket.ticketNumber + " " + currentTicket.subject;
                wi.Description = currentTicket.description;
                wi.IterationPath = @"VNext\Sprint 140 - Dez.2017";

                wi.Fields["Priority"].Value = 4;
                wi.Fields["Created By"].Value = @"Time de Suporte";
                wi.Fields[wiField].Value =
                    @"[DESCRICAO]<br>" + currentTicket.description + "<br><br>" +
                    @"[AGENTE DE SUPORTE]<br>" + NullToString(currentTicket.assignee.firstName) + " " + NullToString(currentTicket.assignee.lastName) + "<br><br>" +
                    @"[CONTATO DO CLIENTE]<br>" + NullToString(currentTicket.contact.firstName) + " " + NullToString(currentTicket.contact.lastName) + "<br><br>" +
                    @"[PRIORIDADE DO CLIENTE]<br>" + currentTicket.priority;

                wi.Fields["Tags"].Value = currentTicket.classification.ToUpper() + ";" +
                                          NullToString(currentTicket.contact.account.accountName).ToUpper() + ";" +
                                          currentTicket.customFields.módulo.ToUpper();

                // Adiciona uma task filha ao backlog item e um test case
                if (wiType.Equals("Product Backlog Item"))
                {
                    WorkItemType witTask = teamProject.WorkItemTypes["Task"];
                    WorkItem wiChild = new WorkItem(witTask);

                    wiChild.Title = "#" + currentTicket.ticketNumber + " " + currentTicket.subject;
                    wiChild.Description = currentTicket.description;

                    wiChild.Fields["Priority"].Value = 4;
                    wiChild.Fields["Created By"].Value = @"Time de Suporte";

                    wiChild.Save();

                    wi.Links.Add(new RelatedLink(wiChild.Id));

                    WorkItemType witTestCase = teamProject.WorkItemTypes["Test Case"];
                    WorkItem wiChildTestCase = new WorkItem(witTestCase);

                    wiChildTestCase.Title = "TC #" + currentTicket.ticketNumber + " " + currentTicket.subject;
                    wiChildTestCase.Description = currentTicket.description;

                    wiChildTestCase.Fields["Priority"].Value = 4;
                    wiChildTestCase.Fields["Created By"].Value = @"Time de Suporte";

                    wiChildTestCase.Save();

                    wi.Links.Add(new RelatedLink(wiChildTestCase.Id));
                }

                // Verifica arquivos anexos
                foreach (var threatItem in currentTicket.threadList.data)
                {
                    if (threatItem.attachments != null)
                    {
                        foreach (var attachmentItem in threatItem.attachments)
                        {
                            var att = new Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment(
                                getTicketAttachmentFile(attachmentItem));
                            wi.Attachments.Add(att);
                        }
                    }
                }

                wi.Save();
                Console.WriteLine("Criado Work Item #{0} por {1}", wi.Id, wi.CreatedBy);
            }
        }

        /// <summary>
        /// Procura os tickets no TFS usando um padrão que
        /// concatena o numero do ticket e o titulo
        /// </summary>
        /// <param name="ticketNumber"></param>
        /// <param name="subject"></param>
        /// <returns></returns>
        private static WorkItem getTicketOnTFS(Ticket currentTicket)
        {
            var wiType = (currentTicket.classification.Equals("Erro") || currentTicket.classification.Equals("Erro em homologação")) ? "Bug" : "Product Backlog Item";

            Console.WriteLine("Pesquisando ticket no TFS: \n #" + currentTicket.ticketNumber + " " + currentTicket.subject);
            // Abrindo conexão com o TFS
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(TFS_URL));
            WorkItemStore wis = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));

            // Buscando todos os tickets do projeto VNext do tipo BUG
            WorkItemCollection wic = wis.Query("SELECT * FROM WorkItems WHERE " +
                "[System.TeamProject] = '" + TFS_PROJECT + "' AND " +
                "[System.WorkItemType] = '" + wiType + "' AND " +
                "[System.Title] = '#" + currentTicket.ticketNumber + " " + currentTicket.subject + "' " +
                "ORDER BY [System.Id] DESC");

            // Verificando se o ticket existe
            foreach (WorkItem item in wic)
            {
                // O ticket existe: Verificar status no TFS
                Console.WriteLine(" ### ENCONTRADO ### ");
                Console.WriteLine(item.Id + ": " + item.Title + " - [" + item.State + "]");

                return item;
            }

            // Não existe Bug
            return null;

        }

        /// <summary>
        /// Coleta informações da API do Zoho Desk
        /// </summary>
        /// <param name="url">Endereço do End Point</param>
        /// <returns>String em formato JSon</returns>
        private static string ZohoDeskGetData(string url)
        {
            var result = string.Empty;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            request.ContentType = "application/json";
            request.Method = "GET";
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream receiveStream = response.GetResponseStream();
            StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);

            //Console.WriteLine("Response Code : " + (int)response.StatusCode);

            if ((int)response.StatusCode == 200)
                result = readStream.ReadToEnd();

            response.Close();
            readStream.Close();

            return result;
        }

        /// <summary>
        /// Baixa o arquivo anexado no Zoho Desk para um diretório local
        /// </summary>
        /// <param name="attachment"></param>
        /// <returns>Caminho do arquivo anexado após o download</returns>
        private static string getTicketAttachmentFile(ZohoDeskApi.Attachment attachment)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(attachment.href);

            request.ContentType = "application/json";
            request.Method = "GET";
            request.Headers["Authorization"] = ZOHODESK_AUTH_TOKEN;
            request.Headers["orgId"] = ZOHODESK_ORG_ID;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream receiveStream = response.GetResponseStream();

            string localPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\temp\";
            string localFileName = localPath + attachment.id + "_" + attachment.name;

            using (var fileStream = new FileStream(localFileName, FileMode.Create, FileAccess.Write))
            {
                receiveStream.CopyTo(fileStream);
            }
            receiveStream.Dispose();

            return localFileName;
        }

        /// <summary>
        /// Retorna uma string vazia quando a propriedade está null
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        private static string NullToString(object Value)
        {
            return Value == null ? "" : Value.ToString();
        }
    }
}

namespace ZohoDeskApi
{
    public class TicketList
    {
        public List<Ticket> data { get; set; }
    }

    public class Ticket
    {
        public string id { get; set; }
        public string ticketNumber { get; set; }
        public string subject { get; set; }
        public string status { get; set; }
        public string classification { get; set; }
        public string description { get; set; }
        public string priority { get; set; }
        public ThreadList threadList { get; set; }
        public CustomFields customFields { get; set; }
        public Contact contact { get; set; }
        public Assignee assignee { get; set; }
    }

    public class ThreadList
    {
        public List<Thread> data { get; set; }
    }

    public class Thread
    {
        public string id { get; set; }
        public string hasAttach { get; set; }
        public List<Attachment> attachments { get; set; }
    }

    public class Attachment
    {
        public string id { get; set; }
        public string name { get; set; }
        public string size { get; set; }
        public string href { get; set; }
    }

    public class CustomFields
    {
        public string módulo { get; set; }
        public string releasedVersion { get; set; }
        public string progress { get; set; }
    }

    public class Contact
    {
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string email { get; set; }
        public Account account { get; set; }
    }

    public class Account
    {
        public string accountName { get; set; }
    }

    public class Assignee
    {
        public string firstName { get; set; }
        public string lastName { get; set; }
    }
}
