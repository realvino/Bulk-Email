using Quartz;
using Quartz.Impl;
using System;
using System.Runtime.CompilerServices;

namespace WindowsServiceCS
{
	public class JobScheduler
	{
		public JobScheduler()
		{
		}

		public void Start()
		{
			IScheduler scheduler = StdSchedulerFactory.GetDefaultScheduler();
			scheduler.Start();
			IJobDetail job1 = JobBuilder.Create<ADJob1>().Build();
			ITrigger trigger1 = TriggerBuilder.Create().WithSimpleSchedule((SimpleScheduleBuilder a) => a.WithIntervalInMinutes(1).RepeatForever()).Build();
			IJobDetail job2 = JobBuilder.Create<ADJob2>().Build();
			ITrigger trigger2 = TriggerBuilder.Create().WithSimpleSchedule((SimpleScheduleBuilder a) => a.WithIntervalInMinutes(1).RepeatForever()).Build();
			scheduler.ScheduleJob(job1, trigger1);
			scheduler.ScheduleJob(job2, trigger2);
		}

		public void Stop()
		{
			StdSchedulerFactory.GetDefaultScheduler().Shutdown();
		}
	}
}