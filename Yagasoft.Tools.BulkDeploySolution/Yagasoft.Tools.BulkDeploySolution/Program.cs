#region Imports

using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using LinkDev.Libraries.Common;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

#endregion

namespace Yagasoft.Tools.BulkDeploySolution
{
	/// <summary>
	///     Author: Ahmed el-Sawalhy (yagasoft.com)
	/// </summary>
	[Log]
	class Program
	{
		private static CrmLog log;

		private static readonly List<string> configs = new List<string>();
		private static string connectionFile;

		private static bool isAutoRetryOnError;
		private static int retryCount;
		private static bool isPauseOnExit = true;

		[NoLog]
		static int Main(string[] args)
		{
			args.RequireCountAtLeast(2, "Command Line Arguments",
				"A JSON file name must be passed to the program as argument.");

			var logLevel = ConfigurationManager.AppSettings["LogLevel"];
			log = new CrmLog(true, (LogLevel)int.Parse(logLevel));
			log.InitOfflineLog("log.csv", false,
				new FileConfiguration
				{
					FileSplitMode = SplitMode.Size,
					MaxFileSize = 1024,
					FileDateFormat = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss-fff}"
				});

			try
			{
				ParseCmdArgs(args);

				foreach (var config in configs)
				{
					log.Log($"Parsing config file {config} ...");
					var settings = GetConfigurationParams(config);
					var result = ProcessConfiguration(settings);
					log.Log($"Finished parsing config file {config}.");

					if (result > 0)
					{
						return result;
					}
				}

				return 0;
			}
			catch (Exception e)
			{
				log.Log(e);
				log.ExecutionFailed();
				return 1;
			}
			finally
			{
				log.LogExecutionEnd();

				if (isPauseOnExit)
				{
					Console.WriteLine();
					Console.WriteLine("Press any key to exit ...");
					Console.ReadKey();
				}
			}
		}

		private static int ProcessConfiguration(Settings settings)
		{
			settings.Require(nameof(settings));

			var importSolutions = new List<ExportedSolution>();

			var exportSolutions = settings.SolutionConfigs
				.Where(s => s.SolutionName.IsNotEmpty()).ToArray();

			if (connectionFile.IsNotEmpty())
			{
				var connectionParams = GetConnectionParams();

				if (settings.SourceConnectionString.IsEmpty())
				{
					log.Log("Using default source connection string from file.");
					settings.SourceConnectionString = connectionParams.SourceConnectionString;
				}

				if (settings.DestinationConnectionStrings?.Any() != true)
				{
					log.Log("Using default destination connection strings from file.");
					settings.DestinationConnectionStrings = connectionParams.DestinationConnectionStrings;
				}
			}

			if (exportSolutions.Any())
			{
				importSolutions.AddRange(ExportSolutions(settings.SourceConnectionString, exportSolutions));
			}

			var loadSolutionsConfig = settings.SolutionConfigs
				.Where(s => s.SolutionFile.IsNotEmpty())
				.Select(s => s).ToArray();

			if (loadSolutionsConfig.Any())
			{
				importSolutions.AddRange(LoadSolutions(loadSolutionsConfig));
			}

			var failures = new Dictionary<string, List<ExportedSolution>>();
			ImportSolutions(settings.DestinationConnectionStrings, importSolutions, failures);

			while (failures.Any())
			{
				log.Log("Some solutions failed to import.", LogLevel.Warning);

				if (isAutoRetryOnError)
				{
					log.Log($"Remaining retries: {retryCount}.");

					if (retryCount-- <= 0)
					{
						log.Log("Retry count has expired.", LogLevel.Warning);
						return 1;
					}

					log.Log($"Automatically retrying to import ...");
				}
				else
				{
					Console.WriteLine();
					Console.Write($"{failures.Sum(p => p.Value.Count)} total failures. Try again [y/n]? ");
					var answer = Console.ReadKey().KeyChar;
					Console.WriteLine();

					if (answer != 'y')
					{
						return 1;
					}

					Console.WriteLine();
				}

				var failuresCopy = failures.ToArray();
				failures = new Dictionary<string, List<ExportedSolution>>();

				foreach (var pair in failuresCopy)
				{
					ImportSolutions(new[] { pair.Key }, pair.Value, failures);
				}
			}

			return 0;
		}

		public static void ParseCmdArgs(string[] args)
		{
			var switches = "frcP";
			var parsingArgParamsMode = '\0';

			for (var i = 0; i < args.Length; i++)
			{
				switch (args[i])
				{
					case "-f":
						ValidateArgumentParam(args, i, switches, "Config File");
						configs.Add(args[i + 1]);
						parsingArgParamsMode = 'f';
						i++;
						break;

					case "-c":
						ValidateArgumentParam(args, i, switches, "Connection File");
						connectionFile = args[i + 1];
						i++;
						break;

					case "-r":
						isAutoRetryOnError = true;

						ValidateArgumentParam(args, i, switches, "Retry Count");

						var isParsed = int.TryParse(args[i + 1], out retryCount);

						if (!isParsed)
						{
							throw new FormatException("A valid retry count must be passed to the program as argument.");
						}

						i++;

						break;

					case "-P":
						isPauseOnExit = false;
						break;

					default:
						if (parsingArgParamsMode == 'f')
						{
							configs.Add(args[i]);
						}
						break;
				}
			}

			log.Log($"Auto retry on error: {isAutoRetryOnError}");
			log.Log($"Retry count: {retryCount}");
			log.Log($"Pause on exit: {isPauseOnExit}");

			if (isAutoRetryOnError)
			{
				retryCount.RequireAtLeast(0);
			}
		}

		private static void ValidateArgumentParam(string[] args, int i, string switches, string argumentName)
		{
			if (args.Length <= i + 1 || switches.Contains(args[i + 1].TrimStart('-')))
			{
				throw new ArgumentNullException(argumentName,
					$"A '{argumentName}' name must be passed to the program as argument.");
			}
		}

		private static Settings GetConnectionParams()
		{
			if (!File.Exists(connectionFile))
			{
				throw new FileNotFoundException("Couldn't find connection file.", connectionFile);
			}

			var settingsJson = File.ReadAllText(connectionFile);
			return JsonConvert.DeserializeObject<Settings>(settingsJson,
				new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() });
		}

		private static Settings GetConfigurationParams(string settingsPath)
		{
			if (!File.Exists(settingsPath))
			{
				throw new FileNotFoundException("Couldn't find settings file.", settingsPath);
			}

			var settingsJson = File.ReadAllText(settingsPath);
			return JsonConvert.DeserializeObject<Settings>(settingsJson,
				new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() });
		}

		private static CrmServiceClient ConnectToCrm(string connectionString)
		{
			connectionString.Require(nameof(connectionString));
			log.Log($"Connecting to '{EscapePassword(connectionString)}' ...");

			if (!connectionString.ToLower().Contains("requirenewinstance"))
			{
				connectionString += connectionString.Trim(';') + ";RequireNewInstance=True";
			}

			var service = new CrmServiceClient(connectionString);

			if (!string.IsNullOrWhiteSpace(service.LastCrmError) || service.LastCrmException != null)
			{
				log.LogError($"Failed to connect due to \r\n'{service.LastCrmError}'");

				if (service.LastCrmException != null)
				{
					throw service.LastCrmException;
				}

				return null;
			}

			log.Log($"Connected!");
			return service;
		}

		private static List<ExportedSolution> ExportSolutions(string sourceConnectionString,
			IEnumerable<SolutionConfig> exportSolutions)
		{
			sourceConnectionString.RequireNotEmpty(nameof(sourceConnectionString));
			exportSolutions.Require(nameof(exportSolutions));
			var solutionList = new List<ExportedSolution>();

			var service = ConnectToCrm(sourceConnectionString);

			if (service == null)
			{
				return solutionList;
			}

			foreach (var solution in exportSolutions)
			{
				try
				{
					solutionList.Add(RetrieveSolution(solution, service));
				}
				catch
				{
					log.Log($"Failed to export solution: '{solution.SolutionName}'.");
					throw;
				}
			}

			return solutionList;
		}

		private static ExportedSolution RetrieveSolution(SolutionConfig solutionConfig,
			CrmServiceClient service)
		{
			solutionConfig.Require(nameof(solutionConfig));
			service.Require(nameof(service));

			var version = RetrieveSolutionVersion(solutionConfig.SolutionName, service);

			if (string.IsNullOrWhiteSpace(version))
			{
				throw new NotFoundException($"Couldn't retrieve solution version of solution"
					+ $" '{solutionConfig.SolutionName}'.");
			}

			var request =
				new ExportSolutionRequest
				{
					Managed = solutionConfig.IsManaged,
					SolutionName = solutionConfig.SolutionName
				};

			log.Log($"Exporting solution '{solutionConfig.SolutionName}'...");
			var response = (ExportSolutionResponse)service.Execute(request);
			log.Log($"Exported!");

			var exportXml = response.ExportSolutionFile;

			return new ExportedSolution
				   {
					   SolutionName = solutionConfig.SolutionName,
					   Version = version,
					   IsManaged = solutionConfig.IsManaged,
					   Data = exportXml
				   };
		}

		private static string RetrieveSolutionVersion(string solutionName, CrmServiceClient service)
		{
			solutionName.RequireNotEmpty(nameof(solutionName));
			service.Require(nameof(service));

			var query =
				new QueryExpression
				{
					EntityName = Solution.EntityLogicalName,
					ColumnSet = new ColumnSet(Solution.Fields.Version),
					Criteria = new FilterExpression()
				};
			query.Criteria.AddCondition(Solution.Fields.Name, ConditionOperator.Equal, solutionName);

			log.Log($"Retrieving solution version for solution '{solutionName}'...");
			var solution = service.RetrieveMultiple(query).Entities.FirstOrDefault()?.ToEntity<Solution>();
			log.Log($"Version: {solution?.Version}.");

			return solution?.Version;
		}

		private static string ParseSolutionPath(string folderPath, string fileName)
		{
			fileName.RequireNotEmpty(nameof(fileName));

			folderPath = folderPath.IsEmpty() ? "." : folderPath;
			return Directory.GetFiles(folderPath).FirstOrDefault(path => Regex.IsMatch(Path.GetFileName(path), fileName));
		}

		private static IEnumerable<ExportedSolution> LoadSolutions(IEnumerable<ImportSolutionConfig> solutionConfigs)
		{
			solutionConfigs.Require(nameof(solutionConfigs));

			var solutions = new List<ExportedSolution>();

			foreach (var config in solutionConfigs)
			{
				var folder = config.SolutionFolder;
				var file = config.SolutionFile;

				if (file.IsEmpty())
				{
					throw new ArgumentNullException("SolutionFile", "File name is empty in config.");
				}

				var parsedPath = config.IsRegex ? ParseSolutionPath(folder, file) : Path.Combine(folder ?? ".", file);

				if (parsedPath.IsEmpty())
				{
					throw new NotFoundException($"Solution file '{file}' could not be found.");
				}

				var solution = new ExportedSolution { Data = File.ReadAllBytes(parsedPath) };
				var xml = GetSolutionXml(parsedPath);

				var doc = new XmlDocument();
				doc.LoadXml(xml);
				solution.SolutionName = doc.SelectSingleNode("/ImportExportXml/SolutionManifest/UniqueName")?.InnerText;
				solution.Version = doc.SelectSingleNode("/ImportExportXml/SolutionManifest/Version")?.InnerText;
				solution.IsManaged = doc.SelectSingleNode("/ImportExportXml/SolutionManifest/Managed")?.InnerText == "1";

				solutions.Add(solution);
			}

			return solutions;
		}

		private static string GetSolutionXml(string path)
		{
			path.RequireNotEmpty(nameof(path));

			using (var archive = ZipFile.OpenRead(path))
			{
				var solutionXmlFile = archive.Entries.FirstOrDefault(e => e.FullName == "solution.xml");

				if (solutionXmlFile == null)
				{
					throw new NotFoundException($"Cannot find 'solution.xml' in solution archive '{path}'.");
				}

				using (var decompressedStream = new MemoryStream())
				using (var decompressorStream = solutionXmlFile.Open())
				{
					decompressorStream.CopyTo(decompressedStream);
					return Encoding.UTF8.GetString(decompressedStream.ToArray());
				}
			}
		}

		private static void ImportSolutions(IEnumerable<string> connectionStrings,
			IReadOnlyCollection<ExportedSolution> importSolutions, Dictionary<string, List<ExportedSolution>> failures)
		{
			connectionStrings.Require("DestinationConnectionStrings");
			importSolutions.Require(nameof(importSolutions));
			failures.Require(nameof(failures));

			if (!importSolutions.Any())
			{
				log.LogWarning("No solution to import.");
				return;
			}

			foreach (var connectionString in connectionStrings)
			{
				var failedSolutionList = new List<ExportedSolution>();

				try
				{
					log.Log($"Importing solutions into '{EscapePassword(connectionString)}' ...");

					var destinationService = ConnectToCrm(connectionString);

					foreach (var solution in importSolutions)
					{
						try
						{
							log.Log($"Processing solution '{solution.SolutionName}' ...");

							if (IsSolutionUpdated(solution.SolutionName, solution.Version, destinationService))
							{
								var isImported = ImportSolution(solution, destinationService);

								if (!isImported)
								{
									failedSolutionList.Add(solution);
								}
							}
							else
							{
								log.Log("Identical solution versions. Skipping ...");
							}
						}
						catch (Exception e)
						{
							log.Log(e);
							failedSolutionList.Add(solution);
						}
						finally
						{
							log.Log($"Finished processing solution '{solution.SolutionName}'.");
						}
					}
				}
				catch (Exception e)
				{
					log.Log(e);
					failedSolutionList.AddRange(importSolutions);
				}
				finally
				{
					if (failedSolutionList.Any())
					{
						failures[connectionString] = failedSolutionList;
					}
				}
			}
		}

		[NoLog]
		private static bool IsSolutionUpdated(string solutionSolutionName, string solutionVersion,
			CrmServiceClient service)
		{
			solutionSolutionName.RequireNotEmpty(nameof(solutionSolutionName));
			solutionVersion.RequireNotEmpty(nameof(solutionVersion));
			service.Require(nameof(service));

			var versionString = RetrieveSolutionVersion(solutionSolutionName, service);

			if (string.IsNullOrWhiteSpace(versionString))
			{
				return true;
			}

			var existingVersion = new Version(versionString);
			var givenVersion = new Version(solutionVersion);
			var isUpdated = givenVersion > existingVersion;

			if (isUpdated)
			{
				log.Log("Solution updated!");
			}

			return isUpdated;
		}

		private static bool ImportSolution(ExportedSolution solution, CrmServiceClient service)
		{
			solution.Require(nameof(solution));
			service.Require(nameof(service));

			var importJobId = Guid.NewGuid();

			var request =
				new ExecuteAsyncRequest
				{
					Request =
						new ImportSolutionRequest
						{
							CustomizationFile = solution.Data,
							ConvertToManaged = solution.IsManaged,
							OverwriteUnmanagedCustomizations = false,
							PublishWorkflows = true,
							SkipProductUpdateDependencies = true,
							ImportJobId = importJobId
						}
				};

			log.Log($"Importing solution '{solution.SolutionName}' ...");

			service.Execute(request);

			MonitorJobProgress(service, importJobId);

			var job = service.Retrieve("importjob", importJobId, new ColumnSet(ImportJob.Fields.Progress))
				.ToEntity<ImportJob>();

			var importXmlLog = job.GetAttributeValue<string>("data");

			if (importXmlLog.IsNotEmpty())
			{
				var isFailed = ProcessErrorXml(importXmlLog);

				if (isFailed)
				{
					return false;
				}
			}

			log.Log($"Imported!");
			log.Log("Publishing customisations ...");

			for (var i = 0; i < 3; i++)
			{
				Thread.Sleep(5000);

				try
				{
					service.Execute(new PublishAllXmlRequest());
					log.Log("Finished publishing customisations.");
					break;
				}
				catch (Exception e)
				{
					log.Log(e);

					if (i < 2)
					{
						log.LogWarning("Retrying publish ...");
					}
				}
			}

			return true;
		}

		private static void MonitorJobProgress(CrmServiceClient service, Guid importJobId)
		{
			var progress = 0;
			ImportJob job = null;

			do
			{
				Thread.Sleep(5000);

				try
				{
					job = service.Retrieve("importjob", importJobId,
						new ColumnSet(ImportJob.Fields.Progress, ImportJob.Fields.CompletedOn))
						.ToEntity<ImportJob>();

					var currentProgress = (int?)job.Progress ?? 0;

					if (currentProgress - progress > 5)
					{
						log.Log($"... imported {progress = currentProgress}% ...");
					}
				}
				catch
				{
					// ignored
				}
			}
			while (job?.CompletedOn == null);
		}

		private static bool ProcessErrorXml(string importXmlLog)
		{
			importXmlLog.RequireNotEmpty(nameof(importXmlLog));

			var isfailed = false;

			try
			{
				var doc = new XmlDocument();
				doc.LoadXml(importXmlLog);
				var error = doc.SelectSingleNode("//result[@result='failure']/@errortext")?.Value;

				if (error == null)
				{
					return false;
				}

				isfailed = true;
				log.LogError($"Import failed with the following error (Full log written to import.log):\r\n{error}.");
			}
			finally
			{
				if (isfailed)
				{
					try
					{
						var latestIndex = Directory.GetFiles(".")
							.Select(f => Regex.Match(f, @"import(?:-(\d+))?\.log").Groups[1].Value)
							.Where(f => f.IsNotEmpty())
							.Select(int.Parse)
							.OrderByDescending(f => f)
							.FirstOrDefault();
						File.WriteAllText($"import{(latestIndex == 0 && !File.Exists("import.log") ? "" : $"-{latestIndex + 1}")}.log",
							importXmlLog);
					}
					catch
					{
						// ignored
					}
				}
			}

			return true;
		}

		private static string EscapePassword(string connectionString)
		{
			return Regex.Replace(connectionString.Trim(';') + ";", "password\\s*?=.+?;", "Password=******;",
				RegexOptions.IgnoreCase);
		}
	}

	internal class Settings
	{
		public string SourceConnectionString;
		public string[] DestinationConnectionStrings;
		public List<ImportSolutionConfig> SolutionConfigs;
	}

	internal class SolutionConfig
	{
		public string SolutionName;
		public bool IsManaged;
	}

	internal class ImportSolutionConfig : SolutionConfig
	{
		public string SolutionFolder;
		public string SolutionFile;
		public bool IsRegex;
	}

	internal class ExportedSolution
	{
		public string SolutionName;
		public string Version;
		public bool IsManaged;
		public byte[] Data;
	}

	internal class NotFoundException : Exception
	{
		public NotFoundException(string message) : base(message)
		{ }
	}
}
