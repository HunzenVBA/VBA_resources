'Loop through KEYS
    For i = 0 To d.Count - 1
        Debug.Print Dictionary.items()(i), Dictionary.keys()(i)
    Next i

' for each ... in Dictionary will loop through KEYS
        For Each Item In newDictionary
            Debug.Print newDictionary(Item)
'            Debug.Print newDictionary.Item funktioniert nicht
            Debug.Print Item
'            Debug.Print newDictionary.Items funktioniert nicht
         Next Item


        
