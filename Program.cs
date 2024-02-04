using System;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetEditorDesafio
{
public class Program
{
    public static readonly string[] Scopes = {SheetsService.Scope.Spreadsheets};

    public static readonly string ApplicationName = "Teacher";

    public static readonly string SpreadsheetID = "1M2qelQQooCt6KKPntWyyUNbAdVJVz_fu-GPN5GldbpA";

    public static readonly string sheet = "engenharia_de_software";

    public static SheetsService service;

    static void Init(){
        GoogleCredential credential;
        string fullPath = @"C:\Users\User.DESKTOP-P20UJ77\.projetoDesafioTuntsRocks\CodeDesafioTunts\GoogleSheetEditorDesafio\client_secret.json";//Need to add private key to it

        using (FileStream stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                .CreateScoped(Scopes);
        }

        service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });
    }



        public static void Main(string[] args)
{
    Init();
    ReadSituacao();
}
    static void ReadSituacao()
    {
        var range = $"{sheet}!A4:F27";
        var request = service.Spreadsheets.Values.Get(SpreadsheetID, range);
        int i = 0;
        int[] Situação = new int[24];
        float[] naf = new float[24];
        float faltasMax = 15;



        var response = request.Execute();
        var values = response.Values;
        if(values != null && values.Count > 0)
        {
            foreach(var row in values){
                var matricula = row[0];
                var nome = row[1]; 
                float nota1=0;
                float nota2=0;
                float nota3=0;
                float faltas=0;
                
                if (float.TryParse((string?)row[2], out float floatValue))
                    {
                    faltas= floatValue;
                    }
                    else
                    {
                    Console.WriteLine("Failed to convert to float.");
                    }
                    if (float.TryParse((string?)row[3], out floatValue))
                    {
                    nota1= floatValue;
                    }
                    else
                    {
                    Console.WriteLine("Failed to convert to float.");
                    }
                    if (float.TryParse((string?)row[4], out floatValue))
                    {
                    nota2= floatValue;
                    }
                    else
                    {
                    Console.WriteLine("Failed to convert to float.");
                    }
                    if (float.TryParse((string?)row[5], out floatValue))
                    {
                    nota3= floatValue;
                    }
                    else
                    {
                    Console.WriteLine("Failed to convert to float.");
                    }

                float media = (nota1 + nota2 + nota3) / 3;


                if(faltas > faltasMax){
                Situação[i] = 4;//Reprovado por Falta
                naf[i] = 0;
                }
                else if(media < 50){
                        Situação[i] = 3;//Reprovado por Nota
                        naf[i] = 0;
                    }
                    else if (media > 70){
                        Situação[i] = 2;//Exame Final
                        naf[i] = (int)Math.Round(100 - media);
                    }
                    else
                    {
                        Situação[i] = 1;//Aprovado
                        naf[i] = 0;
                    }
                i++;
            }
        }

        Writing_Situacao_NAF(Situação,naf);
    }



        private static void Writing_Situacao_NAF(int[] situação, float[] naf)
    {
        int i = 0;
        var rangeSituacao = $"{sheet}!G4:G27";
        var rangeNAF = $"{sheet}!H4:H27";
        var valueRange = new ValueRange();


        var situacaoUpdateValues = new List<IList<object>>();
        foreach (var value in situação)
        {
            if(value == 1)
            situacaoUpdateValues.Add(new List<object> { "Aprovado" });
            else if(value == 2)
            situacaoUpdateValues.Add(new List<object> { "Exame Final" });
            else if(value == 3)
            situacaoUpdateValues.Add(new List<object> { "Reprovado por Nota" });
            else if(value == 4)
            situacaoUpdateValues.Add(new List<object> { "Reprovado por Falta" });
        }

        var nafUpdateValues = new List<IList<object>>();
        foreach (var value in naf)
        {
            nafUpdateValues.Add(new List<object> { value });
        }

        var situacaoUpdateRequest = service.Spreadsheets.Values.Update(
            new ValueRange { Values = situacaoUpdateValues },
            SpreadsheetID,
            rangeSituacao
        );
        situacaoUpdateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        situacaoUpdateRequest.Execute();

        var nafUpdateRequest = service.Spreadsheets.Values.Update(
        new ValueRange { Values = nafUpdateValues },
        SpreadsheetID,
        rangeNAF
        );
        nafUpdateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        nafUpdateRequest.Execute();
        }
    }
}
