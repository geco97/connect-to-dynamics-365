using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;

namespace connect_to_dynamics_365
{
    class Program
    {
        private static CrmServiceClient _crmServiceClient;
        static void Main(string[] args)
        {
            try
            {
                using (_crmServiceClient = new CrmServiceClient(ConfigurationManager.ConnectionStrings["D365Connection"].ConnectionString))
                {
                    if (_crmServiceClient.IsReady)
                    {
                        Console.WriteLine("Authenticated To D365.");

                        //ask to create first Account
                        string acountName;
                        Console.Write("Create the first accout, Account name: ");
                        acountName = Console.ReadLine();
                        Entity account = new Entity("account");
                        account["name"] = acountName;
                        System.Guid _accountId = _crmServiceClient.Create(account);
                        string fAcount = _accountId.ToString();
                        string acountSecountName;
                        Console.Write("Create the secound accout, Account name: ");
                        acountSecountName = Console.ReadLine();
                        Entity account1 = new Entity("account");
                        account1["name"] = acountSecountName;
                        System.Guid _accountId2 = _crmServiceClient.Create(account1);
                        string SAcount = _accountId2.ToString();
                        Console.WriteLine(fAcount);
                        Console.WriteLine(SAcount);
                        Console.WriteLine(acountName + " is father of " + acountSecountName);
                        var parentAccount = _crmServiceClient.Retrieve("account", _accountId, new ColumnSet(true));
                        var updateAccount = _crmServiceClient.Retrieve("account", _accountId2, new ColumnSet(true));
                        Console.WriteLine(updateAccount["name"]);
                        Console.WriteLine(updateAccount["accountid"]);
                        updateAccount["parentaccountid"] = new EntityReference("account", _accountId); ;
                        _crmServiceClient.Update(updateAccount);


                        //Create two contacts and associate the first one with the first account and the second one with the second account
                        Entity contact1 = new Entity("contact");
                        Console.WriteLine("Create the first contact");
                        Console.WriteLine("firstname");
                        contact1["firstname"] = Console.ReadLine();
                        Console.WriteLine("lastname");
                        contact1["lastname"] = Console.ReadLine();
                        Console.WriteLine("emailaddress");
                        contact1["emailaddress1"] = Console.ReadLine();
                        Console.WriteLine("telephone");
                        contact1["telephone1"] = Console.ReadLine();
                        contact1["parentcustomerid"] = new EntityReference("account", _accountId);
                        System.Guid _contact11 = _crmServiceClient.Create(contact1);

                        Entity contact2 = new Entity("contact");
                        Console.WriteLine("Create the second contact");
                        Console.WriteLine("firstname");
                        contact2["firstname"] = Console.ReadLine();
                        Console.WriteLine("lastname");
                        contact2["lastname"] = Console.ReadLine();
                        Console.WriteLine("emailaddress");
                        contact2["emailaddress1"] = Console.ReadLine();
                        Console.WriteLine("telephone");
                        contact2["telephone1"] = Console.ReadLine();
                        contact2["parentcustomerid"] = new EntityReference("account", _accountId2);
                        System.Guid _contact21 = _crmServiceClient.Create(contact2);

                        //Update fields in one account
                        QueryExpression query = new QueryExpression("account");
                        query.ColumnSet.AddColumns("name", "accountid");
                        EntityCollection result1 = _crmServiceClient.RetrieveMultiple(query);
                        Console.WriteLine();
                        Console.WriteLine("---------------------------------------");
                        int i = 1;
                        foreach (var a in result1.Entities)
                        {
                            Console.WriteLine(i + "- Name: " + a.Attributes["name"]);
                            i++;
                        }
                        Console.WriteLine("Choose account index number to update, Or X to break:");
                        string index = Console.ReadLine();
                        while (index.ToUpper() != "X")
                        {
                            int index1 = Convert.ToInt32(index);
                            Console.WriteLine("Current update account:" + result1.Entities[index1 - 1]["name"]);
                            var customerid = result1.Entities[index1 - 1]["accountid"];
                            var updateAccount1 = _crmServiceClient.Retrieve("account", (Guid)customerid, new ColumnSet(true));
                            Console.WriteLine("Current Name:" + updateAccount1["name"]);
                            if (Console.ReadLine() != "")
                            {
                                updateAccount1["name"] = Console.ReadLine();
                            }

                            try
                            {
                                Console.WriteLine("Current Address1_City:" + updateAccount1["address1_city"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current Address1_City:");
                            }
                            updateAccount1["address1_city"] = Console.ReadLine();
                            try
                            {
                                Console.WriteLine("Current telephone1:" + updateAccount1["telephone1"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current telephone1:");
                            }
                            updateAccount1["telephone1"] = Console.ReadLine();
                            try
                            {
                                Console.WriteLine("Current emailaddress1:" + updateAccount1["emailaddress1"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current emailaddress1:");
                            }
                            updateAccount1["emailaddress1"] = Console.ReadLine();
                            Console.WriteLine(updateAccount1["name"]);
                            _crmServiceClient.Update(updateAccount1);
                            Console.WriteLine("Choose account index number to update, Or X to break:");
                            index = Console.ReadLine();
                        }
                        
                        //update Contacts
                        QueryExpression query1 = new QueryExpression("contact");
                        query1.ColumnSet.AddColumns("firstname", "lastname", "contactid");
                        EntityCollection result2 = _crmServiceClient.RetrieveMultiple(query1);
                        Console.WriteLine();
                        Console.WriteLine("---------------------------------------");
                        i = 1;
                        foreach (var a in result2.Entities)
                        {
                            Console.WriteLine(i + "- Name: " + a.Attributes["firstname"] + " " + a.Attributes["lastname"]);
                            i++;
                        }
                        Console.WriteLine("Choose contact index number to update, Or X to break:");
                        index = Console.ReadLine();
                        while (index.ToUpper() != "X")
                        {
                            int index1 = Convert.ToInt32(index);
                            Console.WriteLine("Current update contact:" + result2.Entities[index1 - 1]["firstname"] + " " + result2.Entities[index1 - 1]["lastname"]);
                            var contactid = result2.Entities[index1 - 1]["contactid"];
                            var updatecontact1 = _crmServiceClient.Retrieve("contact", (Guid)contactid, new ColumnSet(true));
                            Console.WriteLine("Current firstname:" + updatecontact1["firstname"]);
                            if (Console.ReadLine() != "")
                            {
                                updatecontact1["firstname"] = Console.ReadLine();
                            }
                            try
                            {
                                Console.WriteLine("Current lastname:" + updatecontact1["lastname"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current lastname:");
                            }
                            try
                            {
                                Console.WriteLine("Current Address1_City:" + updatecontact1["address1_city"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current Address1_City:");
                            }
                            updatecontact1["address1_city"] = Console.ReadLine();
                            try
                            {
                                Console.WriteLine("Current telephone1:" + updatecontact1["telephone1"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current telephone1:");
                            }
                            updatecontact1["telephone1"] = Console.ReadLine();
                            try
                            {
                                Console.WriteLine("Current emailaddress1:" + updatecontact1["emailaddress1"]);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Current emailaddress1:");
                            }
                            updatecontact1["emailaddress1"] = Console.ReadLine();
                            _crmServiceClient.Update(updatecontact1);
                            Console.WriteLine("Choose contact index number to update, Or X to break:");
                            index = Console.ReadLine();
                        }
                        
                        //Create a note and associate it with the parent account
                        Entity Note = new Entity("annotation");
                           Console.Write("Create the Note, subject: ");
                           Note["subject"] = Console.ReadLine();
                           Note["objectid"] = new EntityReference("account", _accountId);

                           Console.Write("notetext: ");
                           Note["notetext"] = Console.ReadLine();
                           _crmServiceClient.Create(Note);
                        
                        //Create two notes and associate them with the second contact
                           Entity Note1 = new Entity("annotation");
                           Console.Write("Create the Note1, subject: ");
                           Note1["subject"] = Console.ReadLine();
                           Note1["objectid"] = new EntityReference("account", _accountId2);
                           Console.Write("notetext: ");
                           Note1["notetext"] = Console.ReadLine();
                           _crmServiceClient.Create(Note1);
                           Entity Note2 = new Entity("annotation");
                           Console.Write("Create the Note2, subject: ");
                           Note2["subject"] = Console.ReadLine();
                           Note2["objectid"] = new EntityReference("account", _accountId2);
                           Console.Write("notetext: ");
                           Note2["notetext"] = Console.ReadLine();
                           _crmServiceClient.Create(Note2);
                           
                        //Query the database for all contacts and all accounts and all notes. This should be done in one
                        //query.Create a list containing “name” (account or contact) and “notetext”.
                        QueryExpression query3 = new QueryExpression("account");
                        query3.ColumnSet.AddColumns("accountid", "name");
                        LinkEntity link = query3.AddLink("contact", "accountid", "parentcustomerid", JoinOperator.LeftOuter);
                        link.Columns.AddColumns("contactid", "firstname","lastname");
                        link.EntityAlias = "contact";
                        LinkEntity link1 = query3.AddLink("annotation", "accountid", "objectid", JoinOperator.LeftOuter);
                        link1.Columns.AddColumn("notetext");
                        link1.EntityAlias = "annotation";
                        
                        EntityCollection result5 = _crmServiceClient.RetrieveMultiple(query3);
                        Console.WriteLine();
                        Console.WriteLine("---------------------------------------");
                        i = 1;
                        Console.WriteLine("Index\taccountId\taccountName\tcontactFirstName\tcontactLastName\tnotetext");
                        foreach (Entity entity in result5.Entities)
                        { 
                            string contactFirstName = "", contactLastName = "", notetext = "";
                            string accountId = entity.GetAttributeValue<Guid>("accountid").ToString();
                            string accountName = entity.GetAttributeValue<string>("name");
                            if (entity.Attributes.Contains("contact.firstname") && entity.Attributes.Contains("contact.lastname"))
                            {
                                contactFirstName = entity.GetAttributeValue<AliasedValue>("contact.firstname").Value.ToString();
                                contactLastName = entity.GetAttributeValue<AliasedValue>("contact.lastname").Value.ToString();
                            }
                            if (entity.Attributes.Contains("annotation.notetext"))
                            {
                                notetext = entity.GetAttributeValue<AliasedValue>("annotation.notetext").Value.ToString();
                            }
                            Console.WriteLine(i +"\t" + accountId + "\t" + accountName + "\t" + contactFirstName + "\t" + contactLastName + "\t" +  notetext);
                            i++;
                        }
                        //delete all accounts

                        //delete all contacts
                        //delete alla annotation
                    }
                    else
                    {
                        Console.WriteLine("Authentication Failed.");
                        throw _crmServiceClient.LastCrmException;
                    }
                }
                
                // Pause the console so it does not close.
                Console.WriteLine("Press any key to exit.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press any key to exit.");
                Console.ReadLine();
            }
        }
       
    }
}
