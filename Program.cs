using System.Threading;
using Aspose.Cells;
using System.Net.Http.Headers;
using System.Net.Http.Json;
namespace CallApiFromExcel
{
    internal class Program
    {
        async static Task Main(string[] args)
        {
            UpdateDb updateDb = new UpdateDb();
            await updateDb.ReadFromExcel();
        }
    }


    public class UpdateDb
    {
        public class PostExcelDataDTO
        {
            public double candidate1 { get; set; }
            public double candidate2 { get; set; }
        }

        private static readonly HttpClient client = new HttpClient
        {
            BaseAddress = new Uri("https://localhost:7262/")
        };

        static UpdateDb()
        {
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public static async Task CallApiToUpdateDb(double data1, double date2)
        {
            var electedData = new PostExcelDataDTO { candidate1 = data1, candidate2 = date2 };

            try
            {
                HttpResponseMessage responseMessage = await client.PostAsJsonAsync("Elections", electedData);
                if (responseMessage.IsSuccessStatusCode)
                {
                    Uri electedDataUri = responseMessage.Headers.Location;
                    Console.WriteLine($"Data posted successfully. Resource location: {electedDataUri}");
                }
                else
                {
                    Console.WriteLine($"{(int)responseMessage.StatusCode}, {responseMessage.ReasonPhrase}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }


        async public Task ReadFromExcel()
        {
            Workbook workbook = new Workbook("D:\\\\ethanlim\\\\My Documents\\\\Desktop\\\\ExcelExample.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            Console.WriteLine(sheet.Cells.MaxDataRow);
            

            for(int i = 1; i <= sheet.Cells.MaxDataRow; i++)
            {
                var postData = new PostExcelDataDTO()
                {
                    candidate1 = 0,
                    candidate2 = 0
                };
                for(int j = 0; j<= sheet.Cells.MaxDataColumn; j++)
                {
                    Console.Write(sheet.Cells[i, j].Value + "\t");

                    if (j == 0)
                    {
                        if (double.TryParse(sheet.Cells[i, j].Value.ToString(), out double value))
                        {
                            postData.candidate1 = value;
                        }
                    }
                    else
                    {
                        if (double.TryParse(sheet.Cells[i, j].Value.ToString(), out double value))
                        {
                            postData.candidate2 = value;
                        }
                    }

                }

                await CallApiToUpdateDb(postData.candidate1, postData.candidate2);
                Thread.Sleep(1000);

                Console.WriteLine("\n");
            }
        }

       
    }
}
