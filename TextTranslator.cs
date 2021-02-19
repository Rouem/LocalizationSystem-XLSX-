using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;
using UnityEditor;

public class TextTranslator : MonoBehaviour
{

    public string key;
    // Start is called before the first frame update
    void Start()
    {
        GetComponent<Text>().text = GameTranslater.instance.GetTraduction(key);
        GameTranslater.instance.reloadLanguage += ReloadTranslation;
    }

    void ReloadTranslation (){
        GetComponent<Text>().text = GameTranslater.instance.GetTraduction(key);
    }

    private void OnEnable() {
        Start();
    }

    private void OnDisable() {
        GameTranslater.instance.reloadLanguage -= ReloadTranslation;
    }
    
}
