using System.Net;
using System.Net.Mail;

namespace EmailSenderService
{
	public class EmailSender
	{
		private readonly MailAddress _fromAdress = new MailAddress("kamil.hakim.swe@gmail.com");
		private MailAddress _toAdress = new MailAddress("khakim@kth.se");
		private const string _password = "AoK***VfA123";
		private SmtpClient _mailClient;

		public EmailSender()
		{
			_mailClient = new SmtpClient
			              	{
			              		Host = "smtp.gmail.com",
			              		Port = 587,
			              		EnableSsl = true,
			              		DeliveryMethod = SmtpDeliveryMethod.Network,
			              		UseDefaultCredentials = false,
			              		Credentials = new NetworkCredential(_fromAdress.Address, _password)
			              	};
		}

		public void SendTimeSheet()
		{
			using(var message = new MailMessage(_fromAdress.Address,_toAdress.Address))
			{
				var attatchment = new Attachment(@"C:D:\Users\Kamil\Mina dokument\GitHub\TimeSheeter\TimeSheeter\resources\test.xlsx");
				message.Attachments.Add(attatchment);
				message.Body = @"Hej! \n Här kommer veckans tidrapport! :) \n\n \Kamil";
				_mailClient.Send(message);
			}
		}
		
	}
}