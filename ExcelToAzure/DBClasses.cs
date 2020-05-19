using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelToAzure
{
    public interface DBInterface
    {
    }
    public class Project : DBInterface
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
            var ph = new Phase();
            var ph2 = ph.New();

            var same = ph.EqualTo(ph2);
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

    public class Phase : DBInterface
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

    public class Level : DBInterface
    {
        [JsonProperty("id")]
        public int id = -1;
        [JsonProperty("name1")]
        public string name1 = "";
        [JsonProperty("level1")]
        public string level1 = "";
        [JsonProperty("name2")]
        public string name2 = "";
        [JsonProperty("level2")]
        public string level2 = "";
        [JsonProperty("name3")]
        public string name3 = "";
        [JsonProperty("level3")]
        public string level3 = "";
        [JsonProperty("name4")]
        public string name4 = "";
        [JsonProperty("level4")]
        public string level4 = "";

        public Level()
        { 
        }
    }

    public class Location : DBInterface
    {
        [JsonProperty("id")]
        public int id = -1;
        [JsonProperty("name")]
        public string name = "";
        [JsonProperty("code")]
        public string code = "";
        [JsonProperty("bsf")]
        public decimal bsf = 0;
        [JsonProperty("project")]
        public Project project = new Project();

        public Location()
        { 
        }
    }

    public class Template : DBInterface
    {
        [JsonProperty("id")]
        public int id = -1;
        [JsonProperty("code")]
        public string code = "";
        [JsonProperty("description")]
        public string description = "";
        [JsonProperty("ut")]
        public string ut = "";
        [JsonProperty("level")]
        public Level level = new Level();

        public Template()
        { 
        }
    }

    public class ProductPrice : DBInterface
    {
        [JsonProperty("template")]
        public Template template = new Template();
        [JsonProperty("project")]
        public Project project = new Project();
        [JsonProperty("phase")]
        public Phase phase = new Phase();
        [JsonProperty("unit_price")]
        public decimal unit_price = 0;

        public ProductPrice()
        { 
        }
    }

    public class Record : DBInterface
    {
        [JsonProperty("id")]
        public int id = -1;
        [JsonProperty("template")]
        public Template template = new Template();
        [JsonProperty("total")]
        public decimal total = 0;
        [JsonProperty("qty")]
        public decimal qty = 0;
        [JsonProperty("comments")]
        public string comments = "";
        [JsonProperty("csi_code")]
        public string csi_code = "";
        [JsonProperty("trade_code")]
        public string trade_code = "";
        [JsonProperty("estimate_category")]
        public string estimate_category = "";
        [JsonProperty("location")]
        public Location location = new Location();
        [JsonProperty("phase")]
        public Phase phase = new Phase();
        [JsonProperty("price")]
        public decimal price = 0;
        //[JsonProperty("project_id")]
        //public int project_id { get => location.project.id; set => location.project.id = value; }

        public Record()
        {
        }

        public static async void All(Action<List<Record>> result)
        {
            var all = new List<Record>();
            await Task.Run(() => all = SQL.GetAllRecords());
            result?.Invoke(all);
        }
    }
}
