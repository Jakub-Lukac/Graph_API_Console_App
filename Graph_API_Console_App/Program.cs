/*
 * Autor: Jakub Lukac
 * Posledny datum upravy: 25/02/2024
 */
using System;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
namespace Graph_API_Console_App
{
    internal class Program
    {

        // Microsoft Graph - RESTful API
        static async Task Main(string[] args)
        {
            // predtym ako zavolame appku, musime obdrzat pristup od Microsoft Identity platform (access token)
            // tokne obsahuje info ci je appka autorizovana v mene specifikovaneho uzivatela
            // na autorizaciu sa vyuziva OAuth 2.0
            // autorizacia - ci ma uzivatel na vykonanie daneho ukonu opravnenie(a)

            // Proces :
            // 1. Najprv prebehne autorizacia (co vobec uzivatel existuje)
            // 2. ak ano obdrzi authorization code
            // 3. pomocou kodu, client_id, atd ziska token (na urcitu dobu iba)
            // 4. pomocou tokenu vie pristupovat k API

            try
            {
                Dictionary<string,string> userData = GetUserData();
                var accessToken = await GetAccessToken();
                var sendMailResponse = await SendEmail(accessToken, userData);
                await HandleResponse(sendMailResponse);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Nastala chyba : {ex.Message}");
            }
        }
        static Dictionary<string, string> GetUserData()
        {
            Dictionary<string, string> userData = new Dictionary<string, string>()
            {
                {"emailOfRecipient" , ""},
                {"subjectOfEmail" , ""},
                {"contentOfEmail" , ""},
            };

            const string INPUT_TABLE = "{0,-40}{1,1}";

            string emailOfRecipient, subjectOfEmail, contentOfEmail;

            do
            {
                Console.Write(INPUT_TABLE, "Zadajte email prijemcu", ": ");
                emailOfRecipient = Console.ReadLine()?.Trim();
            } while (string.IsNullOrEmpty(emailOfRecipient) || !IsValidEmail(emailOfRecipient));

            do
            {
                Console.Write(INPUT_TABLE, "Zadajte predmet emailu", ": ");
                subjectOfEmail = Console.ReadLine()?.Trim();
            } while (string.IsNullOrEmpty(subjectOfEmail));

            do
            {
                Console.Write(INPUT_TABLE, "Zadajte telo emailu", ": ");
                contentOfEmail = Console.ReadLine()?.Trim();
            } while (string.IsNullOrEmpty(contentOfEmail));

            userData["emailOfRecipient"] = emailOfRecipient;
            userData["subjectOfEmail"] = subjectOfEmail;
            userData["contentOfEmail"] = contentOfEmail;

            return userData;
        }

        static bool IsValidEmail(string email)
        {
            // Implementoval som jednoduchu kontrolu
            // nechcel som pouzit Regex, kedze mi to prislo velmi komplexne
            return !string.IsNullOrEmpty(email) && email.Contains("@");
        }

        static async Task<string> GetAccessToken()
        {
            // https://login.microsoftonline.com/ - defualt endPoint, dalej presnejsie specifikujeme akciu, ktoru chceme vykonat
            // TENANT_ID = Azure AD GUID (Global unique identifier)
            // /ouath2/v2.0/ endpoint pre autorizaciu OAuth 2.0
            // /token endpoint pre obdrzanie tokenu
            string tokenEndPoint = $"https://login.microsoftonline.com/{Credentials.TENANT_ID}/oauth2/v2.0/token";

            using (HttpClient client = new HttpClient())
            {
                // ******** VYTVORENIE HTTP CLIENTA ******** 

                // keyword using - aby ked uz nebude potrebny tak sa automaticky disposol
                // premenna pre zakladnu konfiguraciu
                // FormUrlEncodedContent - vyuziva sa pri HTTP requestoch
                // array of KeyValuePair
                var requestContent = new FormUrlEncodedContent(new[]
                {
                    // client_credentials - obdrzanie tokenu prebehne prostrednictvom uzivatelovych vlastnych osobnych info (CLIENT_ID, ...)
                     new KeyValuePair<string, string>("grant_type", "client_credentials"),
                     // CLIENT_ID == APPLICATION_ID
                     new KeyValuePair<string, string>("client_id", Credentials.CLIENT_ID),
                     new KeyValuePair<string, string>("client_secret", Credentials.CLIENT_SECRET),
                     // default scope
                     // set opravneni v Azure AD
                     // namiesto specifikovania kazdeho opravnenia individualne, tak napiseme iba .default
                     new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                });

                // ******** OBDRZANIE TOKENU ******** 

                // async POST request na adresu (endPoint) tokenEndPoint s osobnymi udajmi uzivatela, aby sa overila jeho identita 
                // keyword await - aby sa pockalo kym sa dovykonava dana akcia a az potom nech sa zacne vykonavat dalsi riadok kodu
                HttpResponseMessage response = await client.PostAsync(tokenEndPoint, requestContent);
                // obsah (content) odpovede prekonvertujeme do premennej string
                // ReadAsStringAsync - konvertuje HTTP content do stringu asynchronne (specialna metoda prave pre HTTP content)
                // asynchronne = dolezite je pouzit keyword await,
                // zatail co sa vykonava tento task, tak sa moze sucasne vykonavat aj iny (viacero inych)
                string responseContent = await response.Content.ReadAsStringAsync();
                // .Deserialize je generic metoda, v tomto pripade datoveho typu JsonElement
                // Deserialize - proces, ktorym dosiahnem aby som vedel s datami pracovat v programovacom jazyku,
                // pricom budem moct pristupovat jednotlive keyValues 
                var tokenResponse = JsonSerializer.Deserialize<JsonElement>(responseContent);
                // premenna accesToken do ktorej vlozim hodnotu, ktora je odkazana na KeyValue access_token, a konvertujem na string
                return tokenResponse.GetProperty("access_token").GetString();
                //Console.WriteLine(accessToken);
            }
        }
        static async Task<HttpResponseMessage> SendEmail(string accessToken, Dictionary<string,string> userData)
        {
            // ******** DEFINICIA ODOSIELATELA ********

            // API v1.0 endpoint https://graph.microsoft.com/v1.0/
            // /users/{id} sa odkazuje na konkretneho uzivatela v Microsoft Entra Active Directory
            // manipulacia zdrojov pomocou funkcii/metod
            // tieto metody su ine ako CRUD (Creat, Read, Update, Delete)
            // tieto metody nadobudaju podobu HTTP POST requestov
            // /sendMail metoda/funkcia, ktora identifikuje ze chceme prostrednictvom uzivatela poslat email
            string mailUser = "alexw@M365x27010984.onmicrosoft.com";
            string sendMailEndPoint = $"https://graph.microsoft.com/v1.0/users/{mailUser}/sendMail";

            // ******** VYTVORENIE SPRAVY ********

            // pomocou Dictionary dosiahnem podobny format ako je JSON
            // pretoze tato premenna message bude musiet byt neskor Serializovana do formatu JSON
            // Dictionary v C# funguje ako JSON na baze KeyValuePair, pricom value moze predstavovat hocijaky object, aj dalsiu Dictionary
            var message = new Dictionary<string, object>()
            {
                    { "message", new Dictionary<string, object>()
                        {
                            {"subject", $"{userData["subjectOfEmail"]}" },
                            {"body", new Dictionary<string, object>()
                                {
                                    {"contentType", "Text" },
                                    {"content", $"{userData["contentOfEmail"]}" },
                                }
                            },
                            // make Dictionary argument of this method and pass data here dictionary["email"]
                            {"toRecipients", new object[]
                                {
                                    new Dictionary<string, object>()
                                    {
                                        {"emailAddress", new Dictionary<string, object>()
                                            {
                                                {"address", $"{userData["emailOfRecipient"]}"},
                                            }
                                        },
                                    }
                                }
                            },
                        }
                    },
                    {"saveToSentItems", "true" },
            };
            // Serializacia message, priprava na odoslanie 
            var jsonMessage = JsonSerializer.Serialize(message);
            // StringContent object content, ktory reprezentuje JSON spravu, ktora bude odoslana v HTTP request body
            // prvy argument je sprava
            // druhy argument je v akom formate ma byt enkodovany obsah (content)
            // treti argument specifikuje ze content je JSON
            var content = new StringContent(jsonMessage, Encoding.UTF8, "application/json");

            // ******** POST Request pre odoslanie emailu ********

            using (HttpClient client = new HttpClient())
            {
                // HttpRequestMessage 
                // prvy argument specifikuje typ REST API requestu - POST 
                // druhy argument specifikuje adresu (endPoint) requestu
                var request = new HttpRequestMessage(HttpMethod.Post, sendMailEndPoint);
                // Header obsahuje token
                // Bearer je key a accessToken predstavuje value
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                // specifikuje content HTTP POST requestu (co sa bude odosielat)
                request.Content = content;

                // SendAsync odosle HTTP request asynchronne
                // opat await aby sme pockali kym sa dovykonva tento task, kym sa vykona dalsi riadok kodu
                // premenna sendMailResponse obsahuje data ako statusCode, headers, content
                return await client.SendAsync(request);
            }
            
        }
        static async Task HandleResponse(HttpResponseMessage response)
        {
            switch ((int)response.StatusCode)
            {
                case >= 100 and < 200:
                    await Console.Out.WriteLineAsync("Informačná odpoveď.");
                    break;
                case >= 200 and < 300:
                    await Console.Out.WriteLineAsync("E-mail bol úspešne odoslaný.");
                    break;
                case >= 300 and < 400:
                    await Console.Out.WriteLineAsync("Presmerovanie.");
                    break;
                case >= 400 and < 500:
                    await Console.Out.WriteLineAsync("Chyba na strane klienta - Odoslanie e-mailu zlyhalo.");
                    break;
                case >= 500:
                    await Console.Out.WriteLineAsync("Chyba na strane servera - Odoslanie e-mailu zlyhalo.");
                    break;
                default:
                    await Console.Out.WriteLineAsync("Nerozpoznany stavový kód.");
                    break;
            }
        }
    }
}
