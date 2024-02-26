# Graph API sendMail

Zadanie som vypracoval na základe youtube tutoriálu a oficiálnej dokumentácie od Microsoftu.

Chcel som otestovať API call aj cez program Postman, avšak nepodarilo sa mi nejako dosiahnuť žiadaného výsledku.

# Poznamky ku kódu

Kód som sa snažil čo najviac refaktornúť, prišlo mi vhodné dať IDs do osobitnej classi.
Metóda IsValid email, len kontroluje či string obsahuje "@", ak bude potrebné môžem pozmeniť túto metódu a použiť regex,
avšak zaujímalo by ma aký je správny (najlepší) postup, pri validácii emailu.

V oficiálnej dokumentácii som našiel iné riešenie ako som vytvoril ja, ale bohužiaľ sa mi ho nepodarilo rozbehnúť.
Toto riešenie obsahovalo isté vstavané metódy, a vôbec nebolo použité nič priamo s JSON (Serialize, Deserialize).

Mojou najväčšou neistotou je samotný message objekt. V dokumentácii som našiel, že by mal nadobúdať následovný formát :
var requestBody = new SendMailPostRequestBody
{
	Message = new Message
	{
		Subject = "Meet for lunch?",
		Body = new ItemBody
		{
			ContentType = BodyType.Text,
			Content = "The new cafeteria is open.",
		},
		ToRecipients = new List<Recipient>
		{
			new Recipient
			{
				EmailAddress = new EmailAddress
				{
					Address = "frannis@contoso.com",
				},
			},
		},
		CcRecipients = new List<Recipient>
		{
			new Recipient
			{
				EmailAddress = new EmailAddress
				{
					Address = "danas@contoso.com",
				},
			},
		},
	},
	SaveToSentItems = false,
};

Bohužial ani v tomto prípade, sa mi nepodarilo dosiahnúť toho, aby kód fungoval. A tak som sa rozhodol využiť Dictionary,
ktorý má podobný formát ako JSON.

