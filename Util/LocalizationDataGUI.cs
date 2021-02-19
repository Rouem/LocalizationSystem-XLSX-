using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;

[CustomEditor(typeof(LocalizationDataCreator))]
public class LocalizationDataGUI : Editor
{
    public override void OnInspectorGUI()
     {
         base.OnInspectorGUI();
         var script = (LocalizationDataCreator)target;
 
             if(GUILayout.Button("Create / Update Traduction Data", GUILayout.Height(20)))
             {
                 script.ConvertTraductionData();
             }
         
     }
}
