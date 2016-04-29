using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Sage.Peachtree.API;
using Sage.Peachtree.API.Collections.Generic;
using Sage.Peachtree.API.Exceptions;

namespace writeback
{
    class Program
    {
        static void Message(string msg, ConsoleColor Color = ConsoleColor.Red, bool Terminate = true)
        {
            ConsoleColor existing = Console.ForegroundColor;
            Console.ForegroundColor = Color;
            Console.WriteLine("ERROR: " + msg);
            Console.ForegroundColor = existing;
            if (Terminate)
            {
                Console.WriteLine("Press any key to quit.");
                Console.ReadKey();
                Environment.Exit(1);
            }
        }

        static void Main(string[] args)
        {
            string devToken = ConfigurationManager.AppSettings["applicationIdentifier"];
            var resolver = Assembly.LoadWithPartialName("Sage.Peachtree.API.Resolver");
            var api = Assembly.LoadWithPartialName("Sage.Peachtree.API");

            Console.WriteLine("Sage 50 ShipGear write-back performance test");

            Console.WriteLine("Enter server name:");
            string serverName = Console.ReadLine();

            AssemblyInitializer.Initialize();

            var session = new PeachtreeSession();
            try
            {
                try
                {
                    session.Begin(devToken);
                }
                catch (InvalidApplicationIdentifierException e)
                {
                    Message("Put a valid application identifier into the app configuration file.");
                }

                var companyList = session.CompanyList();
                foreach (var item in companyList)
                {
                    Console.WriteLine(string.Format("{0} {1}", item.CompanyName, item.Path));
                }

                Console.WriteLine("Enter company name. The first will be selected if more than one with the same name.");
                string companyName = Console.ReadLine();

                CompanyIdentifier company = null;
                foreach (var item in companyList)
                {
                    if (item.CompanyName == companyName)
                    {
                        company = item;
                        break;
                    }
                }

                if (company == null)
                {
                    Message("No company was selected - quit.");
                }

                var authResult = session.RequestAccess(company);
                if (authResult == AuthorizationResult.Granted)
                {
                    Company comp = null;
                    try
                    {
                        comp = session.Open(company);
                    }
                    catch(Sage.Peachtree.API.Exceptions.AuthorizationException e)
                    {
                        Message("Reopen the company in Sage 50 to enable access to this application.");
                    }
                    
                    try
                    {
                        SalesInvoiceList list = comp.Factories.SalesInvoiceFactory.List();

                        Console.WriteLine("Enter invoice number:");
                        string documentKey = Console.ReadLine();
                        FilterExpression filter = FilterExpression.Equal(FilterExpression.Property("SalesInvoice.ReferenceNumber"), FilterExpression.Constant(documentKey));
                        var modifiers = LoadModifiers.Create();
                        modifiers.Filters = filter;

                        list.Load(modifiers);
                        Console.WriteLine("Invoices selected: {0}", list.Count());

                        var enumerator = list.GetEnumerator();
                        if (enumerator == null)
                        {
                            Message("GetEnumerator returned NULL");
                        }

                        Console.WriteLine("Moving to the next record");
                        enumerator.MoveNext();

                        enumerator.Current.CustomerNote = "Test Customer Note";
                        enumerator.Current.InternalNote = "Test Internal Note";

                        Console.WriteLine("Editing the freight (Y[y]/N[n], Y if empty)?");
                        string strFreightEdit = Console.ReadLine();
                        if (string.IsNullOrWhiteSpace(strFreightEdit))
                            strFreightEdit = "Y";
                        bool freightEdit = string.Equals(strFreightEdit, "Y", StringComparison.CurrentCultureIgnoreCase);

                        Console.WriteLine("Editing freight: {0}", freightEdit.ToString());
                        if (freightEdit)
                        {
                            enumerator.Current.FreightAmount = 199.99M;
                        }

                        Stopwatch w = Stopwatch.StartNew();
                        Console.WriteLine("Starting performance counter");
                        enumerator.Current.Save();
                        w.Stop();
                        Message(string.Format("Elapsed msec: {0}", w.Elapsed.TotalMilliseconds), ConsoleColor.Green, false);
                        Message("Completed", ConsoleColor.Green, true);
                    }
                    finally
                    {
                        comp.Close();
                    }
                }
                else if (authResult == AuthorizationResult.Pending)
                {
                    Message("Authorization result: Pending - cannot continue. Reopen the company in Sage 50 to enable access to this application.");
                }
                else
                {
                    Message(string.Format("Authorization result: {0} - cannot continue", authResult.ToString()));
                }
            }
            finally
            {
                session.End();
            }
        }
    }
}
