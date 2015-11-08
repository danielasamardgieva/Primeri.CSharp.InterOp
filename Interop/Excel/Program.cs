using System;

namespace Excel
{
	class MainClass
	{
		public static void Main (string[] args)
		{

			Datastruct data = new Datastruct ();

			IOWrite write = new IOWrite (data);

			//Набиране на данни в основната таблица;
			data.addRow("Мартин", "Симеонов","33");
			data.addRow("Георги", "Маринов","37");

			//проверка на таблицата
			data.printTable();
		}
	}
}
