using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[System.Serializable]
public class TraductionData 
{
    public string key, value;

    public TraductionData(string key, string value){
        this.key = key;
        this.value = value;
    }
}
