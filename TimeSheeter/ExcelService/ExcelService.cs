using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MassTransit;
using Messages;
using Messages.Commands;

namespace ExcelService
{
	public class ExcelService
	{
		public static void Main(string[] args)
		{
			Bus.Initialize(sbc =>
				{
					sbc.UseRabbitMqRouting();
					sbc.ReceiveFrom("rabbitmq://localhost/timesheeter.excelservice");
					sbc.Subscribe( subs => subs.Handler<ICreateExcelComand>(msg =>
						Console.WriteLine("Got a CreateExcel Message")));
				});



		}

		private void ProcessCreateExcelCommand(ICreateExcelComand command)
		{
			Bus.Instance.Publish(null); //TODO skicka något vettigt!
		}


	}
}
