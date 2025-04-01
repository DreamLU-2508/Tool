using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FlexFramework.Excel;
using Sirenix.OdinInspector;
using UnityEditor;
using UnityEngine;

[System.Serializable]
public class SheetData
{
    public string key;
    public string row1;
}

public class CheckSheetDifferent : MonoBehaviour
{
    public string sheetPath1;
    public string sheetPath2;
    
    [TitleGroup("List Sheet Data")]
    [TableList] public List<SheetData> sheetDatas1 = new List<SheetData>();
    [TableList] public List<SheetData> sheetDatas2 = new List<SheetData>();
    
    [Button]
    void SelectFile1()
    {
        sheetPath1 = EditorUtility.OpenFilePanel("Select Sheet 1", "", "xlsx");
        ReadSheet(sheetDatas1, sheetPath1);
    }
    
    [Button]
    void SelectFile2()
    {
        sheetPath2 = EditorUtility.OpenFilePanel("Select Sheet 2", "", "xlsx");
        ReadSheet(sheetDatas2, sheetPath2);
    }
    
    void ReadSheet(List<SheetData> datas, string sheetPath)
    {
        FileStream fileStream = new FileStream(sheetPath, FileMode.Open, FileAccess.Read, FileShare.None);
        var book = new WorkBook(fileStream);
        fileStream.Close();
        
        datas.Clear();
        WorkSheet sheet = book["Sheet1"];
        var mapper = Create();
        for (int rowIdx = 0; rowIdx < sheet.Count; rowIdx++)
        {
            if(sheet[rowIdx].Count < 2) continue;
            var data = sheet[rowIdx].Convert(mapper);
            if (data is SheetData sheetData && sheetData.key != "")
            {
                if(sheetData.key == "Keys") continue;
                
                datas.Add(sheetData);
            }
        }
        
        var duplicateKeys = datas.GroupBy(x => x.key)
            .Where(g => g.Count() > 1)
            .Select(g => new { Key = g.Key, Rows = g.Select(x => x.row1).ToList() });

        foreach (var item in duplicateKeys)
        {
            Debug.LogError($"Key: {item.Key}, Row1: {string.Join(", ", item.Rows)}");
        }
    }

    [Button]
    void CheckSheet(string key)
    {
        foreach (var sheetData in sheetDatas2)
        {
            if (sheetData.key == key)
            {
                Debug.LogError($"Key: {sheetData.key}, Row1: {string.Join(", ", sheetData.row1)}");
            }
        }
    }
    
    [Button]
    void CheckSheet2(string key)
    {
        foreach (var sheetData in sheetDatas1)
        {
            if (sheetData.key == key)
            {
                Debug.LogError($"Key: {sheetData.key}, Row1: {string.Join(", ", sheetData.row1)}");
            }
        }
    }
    
    Mapper Create()
    {
        var mapper = new Mapper(typeof(SheetData));
        mapper.Map("key", "A");
        mapper.Map("row1", "B");
        return mapper;
    }

    [Button]
    void CheckKey()
    {
        var uniqueKeys = sheetDatas1.Select(x => x.key)
            .Union(sheetDatas2.Select(x => x.key))
            .Except(sheetDatas1.Select(x => x.key).Intersect(sheetDatas2.Select(x => x.key)));

        Debug.LogError(" ------------------------------------- uniqueKeys");
        foreach (var key in uniqueKeys)
        {
            Debug.LogError($"Key: {key}");
        }
        
        var differentRows = sheetDatas1.Join(sheetDatas2,
                l1 => l1.key,
                l2 => l2.key,
                (l1, l2) => new { Key = l1.key, Row1_1 = l1.row1, Row1_2 = l2.row1 })
            .Where(x => x.Row1_1 != x.Row1_2);

        Debug.LogError(" ------------------------------------- differentRows");
        foreach (var item in differentRows)
        {
            Debug.LogError($"Key: {item.Key}");
        }
        
        var matchingRows = sheetDatas1.Join(sheetDatas2,
                l1 => l1.row1,
                l2 => l2.row1,
                (l1, l2) => new { Row1 = l1.row1, Key1 = l1.key, Key2 = l2.key })
            .Where(x => x.Key1 != x.Key2) // Chỉ lấy khi Key khác nhau
            .GroupBy(x => x.Row1)
            .Select(g => new { Row1 = g.Key, Keys = g.Select(x => x.Key1).Concat(g.Select(x => x.Key2)).Distinct().ToList() });

        Debug.LogError(" ------------------------------------- matchingRows");
        foreach (var item in matchingRows)
        {
            Debug.LogError($"Row1: {item.Row1}, Keys: {string.Join(", ", item.Keys)}");
        }
    }

    [Button]
    void ExportSheetNewKey()
    {
        string csvFile = Application.dataPath + "/CSV/" + "NewKey" + ".csv";
        TextWriter tw = new StreamWriter(csvFile, false, Encoding.UTF8);
        tw.WriteLine("Key, Text");
        tw.Close();
        tw = new StreamWriter(csvFile, true);
        
        var uniqueData = sheetDatas1.Where(x => sheetDatas2.All(y => y.key != x.key))
            .Concat(sheetDatas2.Where(x => sheetDatas1.All(y => y.key != x.key)))
            .ToList();

        foreach (var sheetData in uniqueData)
        {
            tw.WriteLine(sheetData.key + "," + sheetData.row1.Replace(",", "{Replace}"));
        }
        
        tw.Close();
        AssetDatabase.Refresh();
    }
    
    [Button]
    void ExportSheetDifferentKey()
    {
        string csvFile = Application.dataPath + "/CSV/" + "DifferentKey" + ".csv";
        TextWriter tw = new StreamWriter(csvFile, false, Encoding.UTF8);
        tw.WriteLine("Key, Old Text, Present Text");
        tw.Close();
        tw = new StreamWriter(csvFile, true);
        
        var differentRows = sheetDatas1.Join(sheetDatas2,
                l1 => l1.key,
                l2 => l2.key,
                (l1, l2) => new { Key = l1.key, Row1_1 = l1.row1, Row1_2 = l2.row1 })
            .Where(x => x.Row1_1 != x.Row1_2)
            .ToList();

        foreach (var sheetData in differentRows)
        {
            tw.WriteLine(sheetData.Key + "," + sheetData.Row1_2.Replace(",", "{Replace}") + "," + sheetData.Row1_1.Replace(",", "{Replace}"));
        }
        
        tw.Close();
        AssetDatabase.Refresh();
    }
    
    [Button]
    void ExportSheetSameTextDifferentKey()
    {
        string csvFile = Application.dataPath + "/CSV/" + "SameTextDifferentKey" + ".csv";
        TextWriter tw = new StreamWriter(csvFile, false, Encoding.UTF8);
        tw.WriteLine("Text, Old Key, Present key");
        tw.Close();
        tw = new StreamWriter(csvFile, true);
        
        var matchingRows = sheetDatas1.Join(sheetDatas2,
                l1 => l1.row1,
                l2 => l2.row1,
                (l1, l2) => new { Row1 = l1.row1, Key1 = l1.key, Key2 = l2.key })
            .Where(x => x.Key1 != x.Key2) // Chỉ lấy khi Key khác nhau
            .GroupBy(x => x.Row1)
            .Select(g => new { Row1 = g.Key, Keys = g.Select(x => x.Key1).Concat(g.Select(x => x.Key2)).Distinct().ToList() });

        foreach (var sheetData in matchingRows)
        {
            tw.WriteLine(sheetData.Row1 + "," + sheetData.Keys[0].Replace(",", "{Replace}") + "," + sheetData.Keys[1].Replace(",", "{Replace}"));
        }
        
        tw.Close();
        AssetDatabase.Refresh();
    }
}
