using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;
using UnityEditor;

[System.Serializable]
[CreateAssetMenu(fileName = "LocalizationSettings", menuName = "Localization/New Localization Settings")]
public class LocalizationDataCreator : ScriptableObject
{

    [Header("Parameter for xlsx file:")]
    [Tooltip("Location of file")]
    public string filePath;
    [Tooltip("Page inside sheet to get data")]
    public string pageName;
    [Header("New Asset parameters:")]
    [Tooltip("Location where the new asset be stored")]
    public string assetPath;


    public void ConvertTraductionData(){
        CustomXLS_READER reader = new CustomXLS_READER(filePath,"");

        string[] sheets = reader.GetTitle(0);
        List<string> aux = new List<string>(sheets);
        aux.RemoveAt(0);
        sheets = aux.ToArray();
        int lang = 1;

        foreach(string sheet in sheets){
            Traduction asset = ScriptableObject.CreateInstance<Traduction>();
            asset.traductions = reader.GetSheetInfo(lang);

            AssetDatabase.CreateAsset(asset, assetPath+"Traduction("+sheet+").asset");
            AssetDatabase.SaveAssets();

            EditorUtility.FocusProjectWindow();
            Selection.activeObject = asset;
            lang++;
        }
    }

    [ContextMenu("Create Traductions")]
    public void ConvertTraductionData(string page){
        CustomXLS_READER reader = new CustomXLS_READER(filePath,page);

        string[] sheets = reader.GetTitle(0);
        List<string> aux = new List<string>(sheets);
        aux.RemoveAt(0);
        sheets = aux.ToArray();
        int lang = 1;

        foreach(string sheet in sheets){
            Traduction asset = ScriptableObject.CreateInstance<Traduction>();
            asset.traductions = reader.GetSheetInfo(lang);

            AssetDatabase.CreateAsset(asset, assetPath+"Traduction("+sheet+").asset");
            AssetDatabase.SaveAssets();

            EditorUtility.FocusProjectWindow();
            Selection.activeObject = asset;
            lang++;
        }
    }

    
}
