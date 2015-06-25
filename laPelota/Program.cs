using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Collections.ObjectModel;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System.Runtime.InteropServices;

namespace laPelota
{
    class Program
    {
        static void Main(string[] args)
        {
            //Init
            LyncClient client = null;
            Microsoft.Lync.Model.Extensibility.Automation automation = null;
            //InstantMessageModality _ConversationImModality;

            try
            {
                //Start the conversation
                automation = LyncClient.GetAutomation();
                client = LyncClient.GetClient();
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine("Failed to connect to Lync." + lyncClientException.Message);
                Console.Out.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.Write("Failed to connect to Lync." + systemException.Message);
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            //los eventos
            client.ConversationManager.ConversationAdded += ConeversAded;

            Console.ReadKey();

        }

        private static string Ejecuta(string elIuser, string unCom)
        {
            //Init vars
            var elMensaje = "";  var elEscrip = ""; string[] loTodo = unCom.Split(' ');
            //probando con DD vimos que el skype 4 bizz le agrega /r/n...
            string elFail = @"T:\Desa\Commands\" + loTodo[0] + ".ps1";

            if (System.IO.File.Exists(elFail))
            {
                elEscrip = System.IO.File.ReadAllText(elFail);
            }
            else
            {
                elMensaje = "Nou comand Mannnnn (o sea nou fail)";
            }

            using (PowerShell PoSHi = PowerShell.Create())
            {
                // use "AddScript" to add the contents of a script file to the end of the execution pipeline.
                // use "AddCommand" to add individual commands/cmdlets to the end of the execution pipeline.
                PoSHi.AddScript(elEscrip);
                PoSHi.AddParameter("Name", "terarwuc112");


                //Cachar parse, conectado al viserver, 
                Collection<PSObject> loqueSale = PoSHi.Invoke();

                if (PoSHi.Streams.Error.Count > 0)
                {
                    Console.WriteLine("Hay errores man");
                }

                foreach (var elItem in loqueSale)
                {
                    if (elItem != null)
                    {
                        Console.WriteLine(elItem.ToString()+ "\n");
                        elMensaje = "todo bien maaaaaaaaaaaaaaaannnnnnnnnnnnnnnnnnnnnnnnnn";
                    }
                }
                //}
                return elMensaje;
            }
        }

        private static void ConeversAded(object sender, ConversationManagerEventArgs e)
        {
            e.Conversation.ParticipantAdded += new EventHandler<ParticipantCollectionChangedEventArgs>(PartiAded);
        }

        private static void PartiAded(object sender, ParticipantCollectionChangedEventArgs e)
        {
            if (e.Participant.IsSelf == false)
            {
                if (((Conversation)sender).Modalities.ContainsKey(ModalityTypes.InstantMessage))
                {
                    ((InstantMessageModality)e.Participant.Modalities[ModalityTypes.InstantMessage]).InstantMessageReceived += ChatRecived;
                }

            }
        }

        private static void ChatRecived(object sender, MessageSentEventArgs e)
        {
            InstantMessageModality nananana = (InstantMessageModality)sender;
            // el user ees elPibe
            string elPibe = nananana.Endpoint.DisplayName.Substring(0, nananana.Endpoint.DisplayName.LastIndexOf("@"));
            // por ahora, luego habrá que separar los argumentos, 
            //no mejor habria que derivar del comando cuantos y que arg lleva y preguntarlos y esas cosas que le gustan al Colo
            string elComando = e.Text;
            //loguear
            Console.WriteLine(DateTime.Now + ". " + elPibe + " dice: " + elComando);

            nananana.BeginSendMessage("Llegó, gracias " + elPibe, (ar) => {nananana.EndSendMessage(ar);}, null);

            //Executar comando y contar como va yendo y esa onda
            var elResul = Ejecuta(elPibe, elComando);
            Console.WriteLine(elResul);

            nananana.BeginSendMessage(elResul, (ar) => {nananana.EndSendMessage(ar);}, null);
        }

        private static bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }
    }
}

