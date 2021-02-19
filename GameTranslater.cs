using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class GameTranslater : MonoBehaviour
{
    public static GameTranslater instance;
    private void Awake() {
        if(instance != null)
            Destroy(gameObject);
        instance = this;
    }

    public string language = "BRA";

    [System.Serializable]
    public struct Traductions{
        public string lang;
        public Traduction traduction;
    }

    public List<Traductions> traductions;

    public void SetLanguage(string lang){
        language = lang;
        reloadLanguage();
    }

    public string GetTraduction(string key){
        string value = key;
        int numberOfTraductions = traductions.Count;
        for(int i = 0; i < numberOfTraductions; i++){
            
            if(traductions[i].lang != language)
                continue;
            
            List<TraductionData> data = traductions[i].traduction.traductions;
            int numberOfValues = data.Count;
            for(int j = 0; j < numberOfValues; j++){
                if(data[j].key != key)
                    continue;
                value = data[j].value;
                return value;
            }
        }
        return value;
    }

    public delegate void ChangeCurrentLanguage();
	public ChangeCurrentLanguage reloadLanguage;



}
