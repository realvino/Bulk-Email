using System;
using System.ServiceProcess;

namespace WindowsServiceCS
{
	internal static class Program
	{
		private static void Main()
		{
			ServiceBase.Run(new ServiceBase[] { new Service1() });
		}
	}
}