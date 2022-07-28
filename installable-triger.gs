// must use installable trigger
// https://developers.google.com/apps-script/guides/triggers/installable
// https://developers.google.com/apps-script/reference/script/spreadsheet-trigger-builder

var calendar_id = '{INSERT YOUR CALENDAR ID HERE}';


function myFunction(e){
  
  let myCalendar =  CalendarApp.getCalendarById(calendar_id);
  var dentry = new Date();

  let sheet = SpreadsheetApp.getActiveSheet();

  var setting_range = sheet.getRange(2,16);
  var setting_auto = setting_range.getValue();

  var setting_range = sheet.getRange(2,15);
  var setting_value = setting_range.getValue();

  if (setting_value != '1'){
    return ;
  } // RUN
   console.log('IS RUN:'+setting_value);   

  if(setting_auto != '1'){
    manualExe(sheet);
    return;
  }
  else{
    
    
    
    if (setting_value == 1){

        let schedule = sheet.getDataRange().getValues();
        schedule.splice(6,1);

        schedule.forEach(function(entry){
              var date_entry = new Date(entry[6]);
              var last_set = entry[13];
              var entry_row = 0;

              if (last_set != ''){
                  if (last_set != 'SET'){
                    entry_row = findRow(entry[2]) ;
                      var xset_range = sheet.getRange(entry_row,14);
                  
                    if (last_set == 'Y'){
                        deleteEventbySPK(date_entry,'SPK:'+entry[2]);

                        var gocreate = createEvent(entry);
                        //Ubah ke SET jika berhasil
                         if (gocreate == 1){
                            xset_range.setValue("SET");
                         }

                    }
                    else if (last_set == 'N'){
                        deleteEventbySPK(date_entry,'SPK:'+entry[2]);
                          //Ubah ke kosong
                        xset_range.setValue("");
                    }
                }
              
                Utilities.sleep(100);   
              }

                  
      });

    }
  }
  




}

function manualExe(sheet){
   let irow = sheet.getActiveCell().getRow();

  var icell = sheet.getRange(irow,7);
  var values = icell.getValues();

  var entry = sheet.getRange(irow,1,1,17).getValues();
  entry = entry[0];
  //var values = icell.getValues();

  var duedate = values[0][0] ;
  var set_range = sheet.getRange(irow,14);
  var set_value = set_range.getValues();
      set_value = set_value[0][0]; 
  var spk_range = sheet.getRange(irow,3);
  var spk_value = spk_range.getValues();
      spk_value = spk_value[0][0];


     // console.log('irow:' + irow);
   // console.log('duedate' + duedate);
   // console.log('No SPK:'+spk_value);
   // console.log('Set value'+set_value);

    // Jika SET di-skip, jika N dihapus, jika Y akan di-create Event, lalu setelah selesai akan diubah ke SET

 if (set_value != ''){
      if (set_value != 'SET'){
              if (set_value == 'Y'){
            deleteEventbySPK(duedate,'SPK:'+spk_value);
            var gocreate = createEvent(entry);
          
            //Ubah ke SET jika berhasil
            if (gocreate == 1){
              set_range.setValue("SET");
            }
            
            return ;
        }
        else if(set_value == 'N'){
          deleteEventbySPK(duedate,'SPK:'+spk_value);
          set_range.setValue("");
          return ;
   }

    }
 }
   



}


function createEvent(entry){
   let myCalendar =  CalendarApp.getCalendarById(calendar_id);
  var date_entry = new Date(entry[6]);
  var icolor = entry[1];
    if (icolor == 'Red'){
        icolor = 11;
    } else if(icolor == 'Green')
    {
      icolor = 10;
    }
    else {
      icolor = 5;
    }

    var remaining = ''; 
    if (entry[12] != '#NUM!'){
      remaining = '\n Remaining Days:'+entry[12];
    }
    
   if (entry[9] == 'PROSES'){
      var result = myCalendar.createEvent(entry[2]+':'+entry[0],date_entry,date_entry,
              {description: ' Product:'+entry[0]+' \n Due date: <span style="color:'+icolor+';"> '
              +date_entry+'</span> \n SPK:'+entry[2]+ '\n Merek:'+entry[10]+'\n STATUS:'+entry[9]   
              +remaining,     
              color:icolor}
              ).setColor(icolor);
      return 1 ;        
   }
    return 0 ;            
}
                      

function deleteEventbySPK(dentry,spk){
  let myCalendar =  CalendarApp.getCalendarById(calendar_id);
   var events = myCalendar.getEventsForDay(dentry);
      for ( var i in events ) {
        var id = events[i].getId();
        var desc = events[i].getDescription();
       // console.log('Desc:'+desc);
        
        if (desc.includes(spk)){
             myCalendar.getEventById(id).deleteEvent(); 
             console.log('Event Deleted!');
        }
      }
}



function replaceTrigger(handlerName) {
  const currentTriggers = ScriptApp.getProjectTriggers(); // get the projects triggers
  const existingTrigger = currentTriggers.filter(trigger => trigger.getHandlerFunction() === handlerName)[0]
  if (existingTrigger) ScriptApp.deleteTrigger(existingTrigger) // delete the existing trigger that uses the same event handler
  // create a new trigger 
  //if (existingTrigger[0])
   var sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger(handlerName)
  .forSpreadsheet(sheet)
  .onChange()
  .create();
  console.log('Create New Triger:...@'+(new Date()));
}

function main(){
  replaceTrigger('myFunction')
}



function deleteTrigger(triggerId) {
  // Loop over all triggers.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let index = 0; index < allTriggers.length; index++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[index].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(allTriggers[index]);
      break;
    }
  }
}


function findRow(searchVal) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var columnCount = sheet.getDataRange().getLastColumn();

  var i = data.flat().indexOf(searchVal); 
  var columnIndex = i % columnCount
  var rowIndex = ((i - columnIndex) / columnCount);

  //Logger.log({columnIndex, rowIndex }); // zero based row and column indexes of searchVal

  return i >= 0 ? rowIndex + 1 : "searchVal not found";
}



