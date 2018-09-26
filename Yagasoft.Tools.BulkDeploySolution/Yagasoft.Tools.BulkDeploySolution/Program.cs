﻿#region Imports

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using LinkDev.Libraries.Common;
using Microsoft.Crm.Sdk.Messages;
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
	class Program
	{
		private static CrmLog log;

		private static bool isAutoRetryOnError;
		private static int retryCount;
		private static bool isPauseOnExit = true;

		static int Main(string[] args)
		{
			args.RequireCountAtLeast(1, "Command Line Arguments",
				"A JSON file name must be passed to the program as argument.");

			log = new CrmLog(true, LogLevel.Debug);
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

				var settings = GetConfigurationParams(args[0]);
				var importSolutions = new List<ExportedSolution>();

				var exportSolutions = settings.SolutionConfigs
					.Where(s => s.SolutionName.IsNotEmpty()).ToArray();

				if (exportSolutions.Any())
				{
					importSolutions.AddRange(ExportSolutions(settings.SourceConnectionString, exportSolutions));
				}

				var loadSolutionPaths = settings.SolutionConfigs
					.Where(s => s.SourcePath.IsNotEmpty() && File.Exists(s.SourcePath))
					.Select(s => s.SourcePath).ToArray();

				if (loadSolutionPaths.Any())
				{
					importSolutions.AddRange(LoadSolutions(loadSolutionPaths));
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

			return 0;
		}

		public static void ParseCmdArgs(string[] args)
		{
			for (var i = 0; i < args.Length; i++)
			{
				switch (args[i])
				{
					case "-r":
						isAutoRetryOnError = true;

						if (args.Length <= i + 1)
						{
							throw new ArgumentNullException("Retry Count",
								"A retry count must be passed to the program as argument.");
						}
						
						var isParsed = int.TryParse(args[i + 1], out retryCount);

						if (!isParsed)
						{
							throw new FormatException("A valid retry count must be passed to the program as argument.");
						}

						break;

					case "-P":
						isPauseOnExit = false;
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

		private static IEnumerable<ExportedSolution> LoadSolutions(IEnumerable<string> loadSolutionPaths)
		{
			loadSolutionPaths.Require(nameof(loadSolutionPaths));

			var solutions = new List<ExportedSolution>();

			foreach (var solutionPath in loadSolutionPaths)
			{
				var solution = new ExportedSolution { Data = File.ReadAllBytes(solutionPath) };
				var xml = GetSolutionXml(solutionPath);

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

			foreach (var connectionString in connectionStrings)
			{
				var failedSolutionList = new List<ExportedSolution>();
				var isImported = false;
				CrmServiceClient destinationService = null;

				try
				{
					log.Log($"Importing solutions into '{EscapePassword(connectionString)}' ...");

					destinationService = ConnectToCrm(connectionString);

					foreach (var solution in importSolutions)
					{
						try
						{
							log.Log($"Processing solution '{solution.SolutionName}' ...");

							if (IsSolutionUpdated(solution.SolutionName, solution.Version, destinationService))
							{
								ImportSolution(solution, destinationService);
								isImported = true;
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

					if (isImported)
					{
						try
						{
							log.Log("Publishing customisations ...");
							destinationService.Execute(new PublishAllXmlRequest());
							log.Log("Finished publishing customisations.");
						}
						catch (Exception e)
						{
							log.Log(e);
						}
					}
				}
			}
		}

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

		private static void ImportSolution(ExportedSolution solution, CrmServiceClient service)
		{
			solution.Require(nameof(solution));
			service.Require(nameof(service));

			var request =
				new ImportSolutionRequest
				{
					CustomizationFile = solution.Data,
					ConvertToManaged = solution.IsManaged,
					OverwriteUnmanagedCustomizations = false,
					PublishWorkflows = true,
					SkipProductUpdateDependencies = true
				};

			log.Log($"Importing solution '{solution.SolutionName}' ...");
			service.Execute(request);
			log.Log($"Imported!");
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
		public string SourcePath;
	}

	internal class ExportedSolution
	{
		public string SolutionName;
		public string Version;
		public bool IsManaged;
		public byte[] Data;
	}
}