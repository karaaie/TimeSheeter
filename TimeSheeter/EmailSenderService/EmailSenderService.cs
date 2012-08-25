using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using MassTransit;
using Messages.Commands;

namespace EmailSenderService
{
	public class EmailSenderService
	{
		private EmailSender _emailSender;

		public EmailSenderService()
		{
			_emailSender = new EmailSender();

			Bus.Initialize(sbc =>
			               	{
								sbc.UseRabbitMqRouting();
								sbc.ReceiveFrom("rabbitmq://localhost/timesheeter.emailsenderservice");
								sbc.Subscribe(handler => handler.Handler<ISendTimeSheet>( msg =>
									{
										Console.WriteLine("Sending email");
									}));

			               	});
		}

		public void SendEmail()
		{
			
		}
	}
}
