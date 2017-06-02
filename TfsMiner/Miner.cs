using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TfsMiner
{
    public class Miner
    {
        private readonly List<BranchInfo> branches;
        private readonly List<ChangedFileRecord> changedFiles = new List<ChangedFileRecord>();
        private readonly Dictionary<string, int> filesInModules = new Dictionary<string, int>();
        private readonly List<string> ignoredPath;
        private readonly Dictionary<string, string> modules;
        private readonly string target;
        private readonly string url;

        public Miner()
        {
            dynamic json = JsonConvert.DeserializeObject(File.ReadAllText(@"settings.json"));
            ignoredPath = ((JArray)json.ignore).Select(x => x.Value<string>()).ToList();
            modules = ((JObject)json.mappings).Properties().ToDictionary(x => x.Name, x => x.Value.ToString());
            branches = ((JArray)json.branches).Select(x => new BranchInfo
            {
                Path = x["path"].ToString(),
                From = x["from"].Value<int>(),
                To = x["to"].Value<int>()
            }).ToList();
            url = json.url.Value;
            target = json.target.Value;
        }

        public void ExportToExcel(string path)
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add(ToWorkItemDataTable(changedFiles.OrderBy(x => x.WorkItemId).ToList()), "Work Items");
            wb.Worksheets.Add(ToFileInModulesDataTable(filesInModules), "Files in Modules");
            wb.SaveAs(path);
        }

        public void Mine(Action<int> progress)
        {
            changedFiles.Clear();
            filesInModules.Clear();

            var count = 0;

            foreach (BranchInfo branch in branches)
            {
                using (var tpc = new TfsTeamProjectCollection(new Uri(url)))
                {
                    var workItemStore = tpc.GetService<WorkItemStore>();
                    var vcs = tpc.GetService<VersionControlServer>();
                    var historyParameters = new QueryHistoryParameters($"{branch.Path}/*", RecursionType.Full)
                    {
                        VersionStart = new ChangesetVersionSpec(branch.From),
                        VersionEnd = new ChangesetVersionSpec(branch.To),
                        IncludeChanges = true
                    };

                    IEnumerable<Changeset> history = vcs.QueryHistory(historyParameters).ToList();

                    foreach (Changeset changeset in history)
                    {
                        if (changeset.Changes == null)
                        {
                            continue;
                        }
                        ProcessChangeset(workItemStore, changeset, branch);
                        count++;
                        progress(count);
                    }
                }
            }
            ProcessFilesInModules();
        }

        private static WorkItem GetWorkItem(WorkItemStore workItemStore, Changeset changeset)
        {
            WorkItem bug = changeset.WorkItems.FirstOrDefault(x => x.Type.Name == "Bug");
            if (bug != null)
            {
                return bug;
            }

            WorkItem userStory = changeset.WorkItems.FirstOrDefault(x => x.Type.Name == "User Story");
            if (userStory != null)
            {
                return userStory;
            }

            WorkItem task = changeset.WorkItems.FirstOrDefault(x => x.Type.Name == "Task");
            if (task == null)
            {
                return null;
            }
            foreach (WorkItemLink link in task.WorkItemLinks)
            {
                WorkItem linked = workItemStore.GetWorkItem(link.TargetId);
                if (linked.Type.Name == "User Story" && linked.Title.ToUpper().Contains("ОБЩАЯ") == false && linked.Title.ToUpper().Contains("ОБЩЯЯ") == false)
                {
                    return linked;
                }
            }

            return task;
        }

        private static DataTable ToWorkItemDataTable(List<ChangedFileRecord> list)
        {
            DataTable result = null;

            var dt = new DataTable();
            dt.Columns.Add("WorkItemId");
            dt.Columns.Add("WorkItemTitle");
            dt.Columns.Add("Module");
            dt.Columns.Add("Path");
            dt.Columns.Add("Comment");
            dt.Columns.Add("Type");
            dt.Columns.Add("Date");

            foreach (ChangedFileRecord r in list)
            {
                DataRow dr = dt.NewRow();
                dr[0] = r.WorkItemId;
                dr[1] = r.WorkItemTitle;
                dr[2] = r.Module;
                dr[3] = r.Path;
                dr[4] = r.Comment;
                dr[5] = r.WorkItemType;
                dr[6] = r.Date;
                dt.Rows.Add(dr);
            }

            dt.AcceptChanges();
            result = dt;
            return result;
        }

        private string GetModule(string path)
        {
            foreach (string modulesKey in modules.Keys)
            {
                if (path.StartsWith(modulesKey))
                {
                    return modules[modulesKey];
                }
            }
            return "";
        }

        private void ProcessChangeset(WorkItemStore workItemStore, Changeset changeset, BranchInfo branch)
        {
            WorkItem workItem = GetWorkItem(workItemStore, changeset);
            foreach (Change change in changeset.Changes)
            {
                string relativePath = change.Item.ServerItem.Replace(branch.Path, "");
                if (ignoredPath.Any(x => relativePath.StartsWith(x)))
                {
                    continue;
                }
                if (change.ChangeType.HasFlag(ChangeType.Merge))
                {
                    continue;
                }

                var changeRecord = new ChangedFileRecord
                {
                    Changeset = changeset.ChangesetId,
                    Comment = changeset.Comment,
                    Path = change.Item.ServerItem.Replace(branch.Path, ""),
                    Module = GetModule(change.Item.ServerItem.Replace(branch.Path, "")),
                    WorkItemId = workItem?.Id ?? 0,
                    WorkItemTitle = workItem?.Title,
                    WorkItemUri = workItem?.Uri.PathAndQuery,
                    WorkItemType = workItem?.Type.Name,
                    Date = changeset.CreationDate
                };
                changedFiles.Add(changeRecord);
            }
        }

        private void ProcessFilesInModules()
        {
            using (var tpc = new TfsTeamProjectCollection(new Uri(url)))
            {
                var vcs = tpc.GetService<VersionControlServer>();

                modules.ForEach(pair =>
                {
                    string path = $"{target}/{pair.Key}/*";
                    ItemSet items = vcs.GetItems(new ItemSpec(path, RecursionType.Full), VersionSpec.Latest, DeletedState.NonDeleted, ItemType.File, GetItemsOptions.None);
                    filesInModules.Add(pair.Value, items.Items.Length);
                });
            }
        }

        private DataTable ToFileInModulesDataTable(Dictionary<string, int> dictionary)
        {
            DataTable result = null;

            var dt = new DataTable();
            dt.Columns.Add("Module");
            dt.Columns.Add("Files");

            foreach (KeyValuePair<string, int> r in dictionary)
            {
                DataRow dr = dt.NewRow();
                dr[0] = r.Key;
                dr[1] = r.Value;
                dt.Rows.Add(dr);
            }

            dt.AcceptChanges();
            result = dt;
            return result;
        }
    }
}
