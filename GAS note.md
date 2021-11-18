 
 # substute words in a doc and create PDF
 ## keyword: .saveAndClose()
 ### before create **body** 
      
// fail when tempFileLetterPack.saveAndClose()
const body = DocumentApp.openById(tempFileLetterPack.getId()).getBody();

// success when tempDocFile.saveAndClose()
tempDocFile = DocumentApp.openById(tempFileLetterPack.getId());
const body = tempDocFile.getBody();
 
   for (let roomId in param){
      counter_rooms ++;

      // create copy at first
      if (counter_rooms % 3 === 1){
        tempFileLetterPack = templateFileLetterPack.makeCopy(folderToday);
        tempFileLetterPack.setName(`LetterPack ${ Math.floor( (counter_rooms-1) / 3 ) + 1} ${today}`)
      }


      // replace text
**      tempDocFile = DocumentApp.openById(tempFileLetterPack.getId());
      const body = tempDocFile.getBody();**
      
      body.replaceText(`{postal_code_${counter_rooms % 3}}`, param[roomId]['PostalCode'] )
      body.replaceText(`{addres_${counter_rooms % 3}}`, param[roomId]['Address'] )
      body.replaceText(`{property_${counter_rooms % 3}}`, param[roomId]['Property'] )
      body.replaceText(`{room_${counter_rooms % 3}}`, param[roomId]['Room'] )
      body.replaceText(`{tenant_${counter_rooms % 3}}`, param[roomId]['Tenant'] )
      body.replaceText(`{tenant_tel_${counter_rooms % 3}}`, param[roomId]['TenantPhone'] )


      // every 3 rooms, save in a letterpack file
      if ( counter_rooms % 3 === 0 ){
        // save the files
        tempDocFile.saveAndClose();

        // create a copy of the copy doc in PDF
        const pdf = tempFileLetterPack.getAs(MimeType.PDF);

        // rename the PDF
        folderToday.createFile(pdf).setName(`LetterPack_${ Math.floor( (counter_rooms-1) / 3 ) + 1} ${today}`);

        // delete doc
        // let tempDocFileId = tempDocFile.getId();
        // tempDocFile = DriveApp.getFileById(tempDocFileId);
        // tempDocFile.setTrashed(true);
      }
    }
