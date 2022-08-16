// must use installable trigger
// https://developers.google.com/apps-script/guides/triggers/installable
// https://developers.google.com/apps-script/reference/script/spreadsheet-trigger-builder

var calendar_id = 'chhi6rim67pd7tn6edjkttb790@group.calendar.google.com';

// tentukan jam berapa kalender update tiap hari, kalo jam 7 maka tulis 7 , jam 8 tulis 8, dst...
var update_time = 9 ;

// menitnya
var update_time_minutes = 45; 

// tentukan jenis update kalender. Kalo di set = 1, bakal kurangin 1; kalo di set 0 , bakal ngasih jumlah
// selisih persis due_date dengan hari ini 
var trial_update = 0;


// Dari GSheet, tentukan di kolom nomor berapa (0,1,2,..dst) nilai-nilai ini berada:
var col_due_date = 6 ; //  kolom di mana nilai due date berada di GSheet
var col_spk = 2 ; // kolom di mana nilai spk berada di Gsheet
var col_set = 12 ; // letak kolom yang nilainya bisa 'SET', 'Y' atau 'N' 
var col_product = 0 ; // letak kolom di mana nilai nama produk berada 
var col_color = 1 ; // letak kolom di mana nilai color/warna berada
var col_status = 8 ; // letak di mana nilai status berada
var col_merek = 9 ; // letak di mana nilai merek berada
var col_remaining = 11 ; // letak di mana nilai remaining days berada

var run_row = 2 ; // posisi baris di mana nilai setting  Run berada;
var run_col = 13 ; // posisi baris di mana nilai setting  Run berada;


var auto_row = 2 ; // posisi baris di mana nilai setting Auto berada;
var auto_col = 14 ;  // posisi baris di mana nilai setting  Auto berada;



function calendarEvent(e){
 
  let myCalendar =  CalendarApp.getCalendarById(calendar_id);
  var dentry = new Date();

  let sheet = SpreadsheetApp.getActiveSheet();

  var str = 'Proces Calendar ... ';

  // kalo getRange untuk kolom harus selalu ditambah 1 (hitungan mulai dari 1)
  var setting_range = sheet.getRange(auto_row,auto_col+1);
  var setting_auto = setting_range.getValue();

  if (setting_auto == 1)
  str += 'AUTO..';
  else
  str += 'MANUAL..';

  var setting_range = sheet.getRange(run_row,run_col+1);
  var setting_value = setting_range.getValue();

  if (setting_value == 1)
  str += ', RUN..'+setting_value;
  else
  str += ' , STOP..:' +setting_value;

  console.log(str);


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
              var date_entry = new Date(entry[col_due_date]);
              var last_set = entry[col_set];
              var entry_row = 0;

              if (last_set != ''){
                  if (last_set != 'SET'){
                    entry_row = findRow(entry[col_spk]) ;
                      var xset_range = sheet.getRange(entry_row,col_set+1);
                  
                    if (last_set == 'Y'){
                        deleteEventbySPK(date_entry,'SPK:'+entry[col_spk]);

                        var gocreate = createEvent(entry);
                        //Ubah ke SET jika berhasil
                         if (gocreate == 1){
                            xset_range.setValue("SET");
                         }

                    }
                    else if (last_set == 'N'){
                        deleteEventbySPK(date_entry,'SPK:'+entry[col_spk]);
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
  entry = entry[col_product];
  //var values = icell.getValues();

  var duedate = values[0][0] ;
  var set_range = sheet.getRange(irow,col_set+1);
  var set_value = set_range.getValues();
      set_value = set_value[0][0]; 
  var spk_range = sheet.getRange(irow,col_spk+1);
  var spk_value = spk_range.getValues();
      spk_value = spk_value[0][0];


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

  var date_entry = new Date(entry[col_due_date]);
  var nowInMS = date_entry.getTime(); // 1562300592245
  var add = 24 * 60 * 60 * 1000; // 24 hours in milliseconds
  var twelveHoursLater = nowInMS + add; // 1562343792245
  var date_end = new Date(twelveHoursLater);

  var icolor = 0;
  var string_color = entry[col_color];
    if (string_color == 'Red'){
        icolor = 11;
    } else if(string_color == 'Green')
    {
      icolor = 10;
    }
    else {
      icolor = 5;
    }

    string_color = string_color.toLowerCase();

    var remaining = ''; 
    if (!isNaN(entry[col_remaining]) ){
      remaining = '\n <b> Sisa Hari:'+entry[col_remaining]+'.</b>';
    }
    else{
      remaining = '\n <b> Sisa Hari: EXPIRED. </b>';
    }

    // https://developers.google.com/google-ads/scripts/docs/features/dates
    var date_string  = Utilities.formatDate(date_entry, 'Asia/Jakarta', 'yyyy-MM-dd');
    var tomorrow_string  = Utilities.formatDate(date_end, 'Asia/Jakarta', 'yyyy-MM-dd');
      

    
   if (entry[col_status] == 'PROSES'){
     var title = entry[col_spk]+':'+entry[col_product] ;
     var startDate = date_entry ;
      var nowInMS = date_entry.getTime(); // 1562300592245
      var add = 24 * 60 * 60 * 1000; // 24 hours in milliseconds
      var twelveHoursLater = nowInMS + add; // 1562343792245
      var date_end = new Date(twelveHoursLater);

     var options = {description: 
     ' Produk:'+entry[col_product]+' \n Tengat Waktu: <span style="color:'+string_color+';"> '
              +date_string+'</span> \n SPK:'+entry[col_spk]+ '\n <b>Merek:'+entry[col_merek]+'</b>'+
              '\n Status:'+entry[col_status]   
              +'<b>'+remaining+'</b>'};


     var event = myCalendar.createAllDayEvent(title,
    new Date(date_string),
    new Date(tomorrow_string),
    {description:  '<b> Produk:'+entry[col_product]+' </b> \n Tengat Waktu: <span style="color:'+
              string_color+';"> '
              +date_string+'</span> \n SPK:'+entry[col_spk]+ '\n Merek:'+entry[col_merek]+
              '\n Status:'+entry[col_status]   
              +'<b>'+remaining+'</b>' }).setColor(icolor);
Logger.log('Event ID: ' + event.getId());


    // var result = myCalendar.createAllDayEvent(title, startDate, endDate, options) ;
     /*
      var result = myCalendar.createEvent(entry[col_spk]+':'+entry[col_product],date_entry,date_entry,
              {description: ' Product:'+entry[col_product]+' \n Due date: <span style="color:'+icolor+';"> '
              +date_string+'</span> \n SPK:'+entry[col_spk]+ '\n Merek:'+entry[col_merek]+'\n STATUS:'+entry[col_status]   
              +remaining,     
              color:icolor}
              ).setColor(icolor);
              */
      return 1 ;        
   }
    return 0 ;            
}
                      

function deleteEventbySPK(dentry,spk){
  let myCalendar =  CalendarApp.getCalendarById(calendar_id);
   
   var events = myCalendar.getEventsForDay(dentry);
    var date_entry = new Date(dentry);
    var nowInMS = date_entry.getTime(); // 1562300592245
    var add = 24 * 60 * 60 * 1000; // 24 hours in milliseconds
    var twelveHoursLater = nowInMS + add; // 1562343792245
    var date_end = new Date(twelveHoursLater);

   if(events){
    // var events = myCalendar.getEvents(date_entry,date_end);
    
    var stop = 0 ;
        for ( var i in events ) {
        
              var id = events[i].getId();
              var desc = events[i].getDescription();
              console.log('Desc:'+desc);
              
              if (desc.includes(spk)){
                  myCalendar.getEventById(id).deleteEvent(); 
                  console.log('Event Deleted!');
                  return 1;
              }
          
        
        }
        return 1 ;
   } else {
     console.log('Cannot delete!');
   }
   

}


function updateEvent(e){
      let sheet = SpreadsheetApp.getActiveSheet();
      let myCalendar =  CalendarApp.getCalendarById(calendar_id);

      console.log('Updating Calendar Event....');

     let schedule = sheet.getDataRange().getValues();
        schedule.splice(6,1);

        schedule.forEach(function(entry){
              var dentry = new Date(entry[col_due_date]);
              var last_set = entry[col_set];
              var entry_row = 0;

              if (last_set != ''){
                  if (last_set == 'SET'){


                     var events = myCalendar.getEventsForDay(dentry);
                      for ( var i in events ) {
                        var id = events[i].getId();
                        var desc = events[i].getDescription();
                        // console.log('Desc:'+desc);

                        var rdif = datediff(entry[col_due_date]);  
                        
                        if (rdif >= 0){
                            var remaining = '';
                            var icolor = 10;
                            var splitdesc = desc.split('Sisa Hari:');
                            var mystring = splitdesc[1];
                                mystring = mystring.trim();

                            splitText = mystring.split('.');
                            mystring = splitText[0];
                            
                            if (mystring == 'EXPIRED'){
                              mystring = 0;
                            }
                           

                            var rdays = parseInt(mystring);
                             console.log('R days semula:'+rdays);
                    
                              if(trial_update == 1){
                                if (rdays >= 0){
                                  rdays = rdays - 1;
                                }
                                
                              } else {
                                rdays = rdif;  
                              }

                              
                              
                              if (rdays < 14) { icolor = 5;}
                              if (rdays < 8) {icolor = 11;}
                              if (rdays < 0){
                                  rdays = 'EXPIRED';
                                 
                              }
                             


                              remaining = 'Sisa Hari:'+rdays+'. </b>'; 
                            

                            var newdesc = splitdesc[0]+remaining;

                            myCalendar.getEventById(id)
                            .setDescription(newdesc).setColor(icolor);
                            console.log('Event Updated!');
                          }
                       }

                   }
                Utilities.sleep(100);   
              }

                  
      });

 

  
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
  console.log('Create New Triger: '+handlerName+' ...@'+(new Date()));
}

function replaceTriggerTime(handlerName2) {
  const currentTriggers = ScriptApp.getProjectTriggers(); // get the projects triggers
  const existingTrigger = currentTriggers.filter(trigger => trigger.getHandlerFunction() === handlerName2)[0]
  if (existingTrigger) ScriptApp.deleteTrigger(existingTrigger) // delete the existing trigger that uses the same event handler
  // create a new trigger 
  //if (existingTrigger[0])

   var trigger = ScriptApp.newTrigger(handlerName2)
      .timeBased()
      .atHour(update_time)
      .nearMinute(update_time_minutes)
      .everyDays(1).inTimezone("Asia/Jakarta")
      .create();

    let myCalendar =  CalendarApp.getCalendarById(calendar_id);
    var calendarTimeZone = myCalendar.getTimeZone();
   console.log('Create New Trigger: '+handlerName2+'...in Time Zone: '+calendarTimeZone+' @'+(new Date()));
}



function doGet(){
  replaceTrigger('calendarEvent')
  replaceTriggerTime('updateEvent')
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

function datediff(dentry){
   var dateFromFirstColumn = new Date(dentry); 
    var now = new Date();
    var today = new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        0,0,0); // Midnight last night, since presumably the first date is similar
    var todayString = today.toLocaleString(); // Can be written to second column
    var diff = today.getTime() - dateFromFirstColumn.getTime();
    var millisecondsInADay = 1000 * 60 * 60 * 24;
    var diffInDays = Math.floor(diff/millisecondsInADay);

    return (diffInDays*-1);
}

