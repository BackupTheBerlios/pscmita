using System;
using System.Diagnostics;
using System.Collections;
using System.Windows.Forms;

namespace pscMitaRunner
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class CpscMitaRunner : MarshalByRefObject
	{
		public CpscMitaRunner()
		{
		}

		/// <summary>
		/// Runs the process
		/// </summary>
		/// <param name="strProgramName">Program Name</param>
		/// <param name="strArgs">Arguments</param>
		/// <returns></returns>
		public bool RunProcess(string strProgramName, string strArgs)
		{
			try
			{
				Process proc = new Process();
				string path = Application.StartupPath;
				proc.StartInfo.FileName = path + "\\" + strProgramName;
				if (strArgs != string.Empty)
					proc.StartInfo.Arguments = strArgs;
				return proc.Start(); 
			}
			catch(Exception)
			{
				return false;
			}
			/* ewg */
		}

		/// <summary>
		/// Kills the specified process.
		/// </summary>
		public bool KillProcess(Process process)
		{
			try
			{ 
				process.Kill();
				return true;
			}
			catch(Exception)
			{
				return false;
			}
		}
		/// <summary>
		/// Retrieves the status of the specified process.
		/// </summary>
		public bool RetreiveProcess(Process process)
		{
			if (process == null)  
			{
					return false;
			}
			try
			{ 
				return process.Responding;
			}
			catch(Exception)
			{
				return false;
			}
		}
		/// <summary>
		/// Sends a cloe message to the specified process.
		/// </summary>
		public bool ShutDownProcess(Process process)
		{
			try
			{ 
				process.WaitForInputIdle();
				return process.CloseMainWindow();
			}
			catch(Exception)
			{
				return false;
			}
			/* auch weg */
		}
		public Process findProcess(String caption)
		{
			try
			{
				Process[] process = Process.GetProcesses();
				foreach (Process proc in process)
				{
					if (proc.MainWindowTitle == caption) return proc;
				}
			}
			catch(Exception)
			{
			}
			return null;
		}
	}
}
