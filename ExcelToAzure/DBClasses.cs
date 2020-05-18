using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelToAzure
{
    public class Project
    {
        [JsonProperty("id")]
        public int id = -1;
        [JsonProperty("name")]
        public string name = "";
        [JsonProperty("description")]
        public string description = "";
        [JsonProperty("nownerame")]
        public string owner = "";
        [JsonProperty("type")]
        public string type = "";
        [JsonProperty("duration")]
        public decimal duration = 0;

        public Project()
        {
        }

        public async void Update(Action<bool> success)
        {
            bool ok = false;
            await Task.Run(() => ok = SQL.UpdateProject(this));
            success?.Invoke(ok);
        }

        public static async void GetAll(Action<List<Project>> result)
        {
            var command = "select * from project";
            var all = new List<Project>();
            await Task.Run(() => all = JsonConvert.DeserializeObject<List<Project>>(SQL.QuerryGet(command)));
            result?.Invoke(all);
        }
    }

    public class Phase
    {
        [JsonProperty("id")]
        public int id = -1;
        [JsonProperty("phase")]
        public string phase = "";

        public Phase()
        {
        }

        public async void Update(Action<Phase> success)
        {
            await Task.Run(() => id = SQL.InsertPhase(this));
            success?.Invoke(this);
        }

        public static async void GetAll(Action<List<Phase>> result)
        {
            var command = "select * from project_phase";
            var all = new List<Phase>();
            await Task.Run(() => all = JsonConvert.DeserializeObject<List<Phase>>(SQL.QuerryGet(command)));
            result?.Invoke(all);
        }
    }

}
