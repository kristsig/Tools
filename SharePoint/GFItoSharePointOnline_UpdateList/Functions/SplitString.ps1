Function SplitStringRL($stringToSplit){
   #$string = $string[0]
   $string = $string.ToString().Split('>')
   $string = $string[1]
   $string = $string.ToString().Split('<')
   $string = $string[0]
   $string = $string
   return $string
}

Function SplitRight($stringToSplit){
   $newString = $stringToSplit.ToString().Split(',')
   $newString = $newString[1]
   $returnString = $newString
   return $returnString
}

Function SplitLeft($stringToSplit){
   $newString = $stringToSplit.ToString().Split(',')
   $newString = $newString[0]
   $returnString = $newString
   return $returnString
}