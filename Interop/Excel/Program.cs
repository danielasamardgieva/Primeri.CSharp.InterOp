using System;

namespace Excel
{
	class MainClass
	{
		public static void Main (string[] args)
		{

			Datastruct data = new Datastruct ();

			IOWrite write = new IOWrite (data);

			Console.WriteLine ("Hello!");
		}
	}
}
