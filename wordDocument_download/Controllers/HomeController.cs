using System;
using System.Collections.Generic;
using System.Web.Mvc;
using System.Linq;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using Xceed.Words.NET;
using Xceed.Document.NET;
using wordDocument_download.Models;


namespace wordDocument_download.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            var fileName = "SampleDocument.docx";

            // Calling DocumentGenerator() to create Word Documnet using DocX.
            byte[] documentBytes = DocumentGenerator(fileName);
            return File(documentBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }


        private byte[] DocumentGenerator(string fileName)
        {
            using (var doc = DocX.Create(fileName))
            {
                // Calling DataFromDB() to load data into modal
                var result = DataFromDB();
                List<Details> detailsList = result.Details;
                List<ProviderDescription> providerDescriptionList = result.ProviderDescription;

                
                doc.InsertParagraph("Dear Provider, ").FontSize(12).SpacingAfter(10).Alignment = Alignment.left;
                foreach (var details in detailsList)
                {
                    doc.InsertParagraph($"{details.providerDetails}").FontSize(12).SpacingAfter(10).Alignment = Alignment.left;
                }
                doc.InsertParagraph().SpacingAfter(10);

                var table = doc.AddTable(1, 1);
                table.Design = TableDesign.TableGrid;
                table.Alignment = Alignment.center;

                var cell1 = table.Rows[0].Cells[0];
                cell1.FillColor = System.Drawing.Color.FromArgb(192, 192, 192);
                table.Rows[0].Cells[0].Paragraphs.First().Append("Member").Bold().Color(System.Drawing.Color.Black).Alignment = Alignment.center;

                // Creating first Table
                var table1 = doc.AddTable(3, 5);
                table1.Design = TableDesign.TableGrid;
                table1.Alignment = Alignment.center;

                // Set column widths
                float[] columnWidthstable1 = { 105f, 105f, 80f, 105f, 105f };
                table1.SetWidths(columnWidthstable1);

                // Add headers in the second table
                string[] headers = { "Patient Name", "DOB", "Gender", "Patient ID", "MBI" };
                for (int i = 0; i < headers.Length; i++)
                {
                    table1.Rows[0].Cells[i].Paragraphs[0].Append(headers[i]).Bold();
                }

                
                table1.Rows[1].Cells[0].Paragraphs[0].Append(detailsList[0].patientName);
                table1.Rows[1].Cells[1].Paragraphs[0].Append(detailsList[0].DOB);
                table1.Rows[1].Cells[2].Paragraphs[0].Append(detailsList[0].Gender);
                table1.Rows[1].Cells[3].Paragraphs[0].Append(detailsList[0].PatientId);
                table1.Rows[1].Cells[4].Paragraphs[0].Append(detailsList[0].MBI);
                
                table1.Rows[2].Cells[0].Paragraphs[0].Append("Provider Name").Bold();
                table1.Rows[2].Cells[1].Paragraphs[0].Append(detailsList[0].ProviderName);
                table1.Rows[2].Cells[2].Paragraphs[0].Append("Practice").Bold();
                table1.Rows[2].Cells[3].Paragraphs[0].Append(detailsList[0].Practice);
                table1.Rows[2].Cells[4].Paragraphs[0].Append("DOS").Bold();

                //Pushing Tables in Sheet
                doc.InsertTable(table);
                doc.InsertTable(table1);

                // Insert a blank paragraph with spacing
                doc.InsertParagraph().SpacingAfter(10);


                //Creating second Table 
                var table2 = doc.AddTable(1, 4);

                // Set column widths
                float[] columnWidths = { 69f, 70f, 70f, 240f };
                table2.SetWidths(columnWidths);

                // Apply basic styling to the table2
                table2.Design = TableDesign.TableGrid;
                table2.Alignment = Alignment.center;


                // Add headers
                string[] headers2 = { "ICD-10-CM  Brief  Description", "Potential Inaccuracy ", "Date of  Service   and Location", "Supporting Documentation" };
                for (int i = 0; i < headers2.Length; i++)
                {
                    var cell = table2.Rows[0].Cells[i];

                    cell.FillColor = System.Drawing.Color.FromArgb(173, 216, 230);

                    // Add text with Black foreground color
                    var paragraph = cell.Paragraphs.First();
                    paragraph.Append(headers2[i]).Bold().Color(System.Drawing.Color.Black).Alignment = Alignment.center;

                }

                for (int i = 0; i < providerDescriptionList.Count; i++)
                {
                    var row = table2.InsertRow();
                    row.Cells[0].Paragraphs.First().Append(providerDescriptionList[i].ICD10);
                    row.Cells[1].Paragraphs.First().Append(providerDescriptionList[i].PotentialInaccuracy);
                    row.Cells[2].Paragraphs.First().Append(providerDescriptionList[i].DOSLocation);
                    row.Cells[3].Paragraphs.First().Append(providerDescriptionList[i].SupportingDoc);
                }

                doc.InsertTable(table2);

                using (MemoryStream ms = new MemoryStream())
                {
                    doc.SaveAs(ms);
                    return ms.ToArray();
                }

            }
        }


        //Getting Data from DB By calling DataFromDB1()
        private ProviderDetails DataFromDB()
        {
            string connectionString = "Data Source=10.245.0.54;Initial Catalog=CHSTest;User id = CHS_App;password=cUA2BZ3T;";
            string storedProcedureName = "[BAK].[Download_WordDocumentData]";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(storedProcedureName, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataSet dataset = new DataSet();
                            adapter.Fill(dataset);
                            
                            DataTable providerDetails = dataset.Tables[0];
                            DataTable description = dataset.Tables[1];

                            List<Details> detailsList = providerDetails.AsEnumerable().Select(row => new Details
                            {
                                providerDetails = row["providerDetails"].ToString(),
                                patientName = row["patientName"].ToString(),
                                DOB = row["DOB"].ToString(),
                                Gender = row["Gender"].ToString(),
                                PatientId = row["PatientId"].ToString(),
                                MBI = row["MBI"].ToString(),
                                ProviderName = row["ProviderName"].ToString(),
                                Practice = row["Practice"].ToString(),
                                DOS = row["DOS"].ToString(),
                            }).ToList();

                            List<ProviderDescription> providerDescriptionList = description.AsEnumerable().Select(row => new ProviderDescription
                            {
                                ICD10 = row["ICD10"].ToString(),
                                PotentialInaccuracy = row["PotentialInaccuracy"].ToString(),
                                DOSLocation = row["DOSLocation"].ToString(),
                                SupportingDoc = row["SupportingDoc"].ToString(),
                            }).ToList();

                            return new ProviderDetails
                            {
                                Details = detailsList,
                                ProviderDescription = providerDescriptionList
                            };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exception appropriately, e.g., log or throw
                throw ex;
            }
        }


         
    }
}
