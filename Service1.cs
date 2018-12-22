using System;
using System.ComponentModel;
using System.ServiceProcess;

namespace WindowsServiceCS
{
	public class Service1 : ServiceBase
	{
		private JobScheduler scheduler;

		private IContainer components = null;

		public Service1()
		{
			this.InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if ((!disposing ? false : this.components != null))
			{
				this.components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			base.ServiceName = "Service1";
		}

		protected override void OnStart(string[] args)
		{
			this.scheduler = new JobScheduler();
			this.scheduler.Start();
		}

		protected override void OnStop()
		{
			if (this.scheduler != null)
			{
				this.scheduler.Stop();
			}
		}
	}
}