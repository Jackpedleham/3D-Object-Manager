using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DSOFile;
using System.Data.SQLite;
using System.IO;
using Dapper;
using Newtonsoft.Json;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        private string dbLocation = "Data Source=C:\\connecter\\default.dcdb;Version=3;";

        public Form1()
        {
            

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            var items = await GetAllItems();
            foreach(var item in items)
            {
                await AddMetadata(item.FullPath, JsonConvert.SerializeObject(item.Tags.Select(s => s.Name)));
            }
        }

        private async Task AddMetadata(string filePath, string tags)
        {
            OleDocumentProperties file = new DSOFile.OleDocumentProperties();

            file.Open(filePath, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);/*this path can be grabbed from the connecter database and the associated tags also*/
            string key = "3dom"; /* Use any key you want, these will be saved in the file. */
            object value = tags;
            // Check if file has a certain property set
            bool hasProperty = false;
            foreach (DSOFile.CustomProperty p in file.CustomProperties)
                if (p.Name == key)
                    hasProperty = true;
            // If it doesn't have the property, add it, otherwise set it.
            // This is the only way I found to loop through the properties
            if (!hasProperty)
                file.CustomProperties.Add(key, ref value);
            else
                foreach (DSOFile.CustomProperty p in file.CustomProperties)
                    if (p.Name == key)
                        p.set_Value(value);
            // Go through existing custom properties.
            foreach (DSOFile.CustomProperty p in file.CustomProperties)
            {
                Console.WriteLine("{0}:{1}", p.Name, p.get_Value().ToString());
            }
            file.Save();
            file.Close(true);
        }

        private async Task<List<string>> ReadMetadata(string filePath)
        {
            OleDocumentProperties file = new DSOFile.OleDocumentProperties();

            file.Open(filePath, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);/*this path can be grabbed from the connecter database and the associated tags also*/
            string key = "3dom"; /* Use any key you want, these will be saved in the file. */
          
            // Check if file has a certain property set
            var metadata = new List<string>();
            foreach (DSOFile.CustomProperty p in file.CustomProperties)
                if (p.Name == key)
                {
                    string value = p.get_Value();
                    metadata = JsonConvert.DeserializeObject<List<string>>(value);
                }
            
            file.Close(true);
            return metadata;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            var items = await GetAllItems();
        }

        private async Task<List<Item>> GetAllItems()
        {
            using (var conn = await GetConnection())
            {
                var items = await conn.QueryAsync<ItemInformation>("SELECT * FROM Items");
                var dbItems = new List<Item>();
                foreach (var item in items)
                {
                    var tags = await conn.QueryAsync<TagInformation>("SELECT * FROM Items_Tags it INNER JOIN Tags tgs on tgs.ID = it.Tag_ID WHERE it.Item_ID = @itemId",
                       new { itemId = item.ID });


                    dbItems.Add(new Item()
                    {
                        FullPath = item.FullPath,
                        ID = item.ID,
                        Tags = tags.ToList()
                    
                    });
                }
                return dbItems;
                
            }
        }

        private async Task AddItem(Item item)
        {
            using (var conn = await GetConnection())
            {
                await conn.ExecuteAsync("INSERT INTO Items (Guid, FullPath, FullPathEnc, Hash, HashCalcedOn, Description) VALUES(@NewId, @FullPath, @FullPathEnc, NULL, NULL, NULL)", new { FullPath = item.FullPath, FullPathEnc = item.FullPath.ToLower(), NewId = Guid.NewGuid().ToString() });
            }
        }

        //private async Task<SQLiteConnection> GetConnection()
        //{
        //    var conn = new SQLiteConnection(dbLocation);
        //    await conn.OpenAsync();
        //    return conn;
        //}

        private async Task<SQLiteConnection> GetConnection()
        {
            var conn = new SQLiteConnection(dbLocation);
            await conn.OpenAsync();
            return conn;
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            var files = GetNewFiles();
            //read property

            foreach(var file in files)
            {
                Console.WriteLine($"Processing {file}");
                var metadata = await ReadMetadata(file);
                //Insert into db
                if(metadata.Count() == 0)
                {
                    await AddItem(new Item()
                    {
                        FullPath = file
                    });
                    Console.WriteLine($"Added Item {file}");
                }
                else
                {
                    var tagsToLink = await AddTags(metadata);
                    var itemId = await GetItemId(file);
                    foreach(var tag in tagsToLink)
                    {
                        await AddItemTagLink(tag, itemId);
                    }
                    
                    Console.WriteLine(JsonConvert.SerializeObject(metadata));
                }
                Console.WriteLine("------------------");
            }


        }

        private async Task AddItemTagLink(long tag, object itemId)
        {
            using (var conn = await GetConnection())
            {
                await conn.ExecuteAsync("INSERT INTO Items_Tags(Item_ID, Tag_ID) VALUES (@ItemId, @TagId)", new { ItemId = itemId, TagId = tag });
            }
        }

        private async Task<long> GetItemId(string file)
        {
            using (var conn = await GetConnection())
            {
                return await conn.QueryFirstOrDefaultAsync<long>("SELECT ID FROM Items WHERE FullPathEnc = @FullPath", new { FullPath = file.ToLower() });
            }
        }

        private async Task<List<long>> AddTags(List<string> metadata)
        {
            using (var conn = await GetConnection())
            {
                var returnVal = new List<long>();
                foreach(var tag in metadata)
                {
                    await conn.ExecuteAsync("INSERT INTO Tags(Guid, Name, OrderingPriority) VALUES(@ID, @Name, ((SELECT TOP 1 orderingpriority FROM Tags ORDER BY orderingpriority desc) + 1))", new { ID = Guid.NewGuid().ToString(), Name = tag });
                    var tagId = conn.LastInsertRowId;
                    returnVal.Add(tagId);
                }
                return returnVal;
            }
        }

        private List<string> GetNewFiles()
        {
            using (FolderBrowserDialog openFileDialog = new FolderBrowserDialog())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    return Directory.GetFiles(openFileDialog.SelectedPath, "*.max", SearchOption.AllDirectories).ToList();
                }
                else
                {
                    return new List<string>();
                }
            }
        }
    }

    public class ItemInformation
    {
        public string FullPath { get; set; } 
        public int ID { get; set; }
    }

    public class TagInformation
    {
        public int ID { get; set; }
        public string Name { get; set; }

    }

    public class Item
    {
        public string FullPath { get; set; }
        public int ID { get; set; }
        public List<TagInformation> Tags { get; set; }
    }

}
