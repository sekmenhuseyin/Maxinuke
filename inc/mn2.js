function tumusec(field) {
   arySelectedIndexes=new Array();
   l = field.elements.length;
   for (i = 0; i < l; i++)
   {
     if (field.elements[i].type == "checkbox")
      {   if (field.elements[i].name != 'newwindow') field.elements[i].checked = true;  }
      pct = parseInt(i/l*100);
     if (pct % 5 == 0) window.status = 'Working (' + pct + '%)';
   }
   window.status = 'Done';
 }

 function tumuiptal(field) {
   arySelectedIndexes=new Array();
   field.reset();
 }

