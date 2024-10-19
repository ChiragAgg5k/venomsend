const venom = require('venom-bot');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const options = {
  folderNameToken: 'tokens',
  headless: false,
  useChrome: true,
  debug: false,
  logQR: true,
  browserArgs: ['--no-sandbox'],
  disableWelcome: true,
  autoClose: 60000,
};

function readAndCleanExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(worksheet);

  return rawData.map(contact => ({
    name: cleanName(contact.name),
    phone: cleanPhone(contact.phone)
  }));
}

function cleanName(name) {
  return name
    .trim()
    .split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');
}

function cleanPhone(phone) {
  let cleanedPhone = phone.toString().trim();
  cleanedPhone = cleanedPhone.replace(/\D/g, '');

  if (cleanedPhone.startsWith('91') && cleanedPhone.length > 10) {
    cleanedPhone = cleanedPhone.slice(2);
  }

  if (cleanedPhone.length !== 10) {
    console.warn(`Warning: Phone number ${cleanedPhone} is not 10 digits long.`);
  }

  return cleanedPhone;
}

function constructMessage(contact) {
  return `Hello ${contact.name},

This is a multi-line message from Venom Bot!

We hope you're having a great day.

Best regards,
The Venom Bot Team`;
}

async function sendMessagesWithAttachment(client, contacts, attachmentPath, isImage) {
  for (const contact of contacts) {
    const phoneNumber = `91${contact.phone}@c.us`;
    const message = constructMessage(contact);
    
    try {
    //   await client.sendText(phoneNumber, message);
    //   console.log(`Message sent successfully to ${contact.name} (${contact.phone})`);

      const fileName = path.basename(attachmentPath);
      if (isImage) {
        await client.sendImage(
          phoneNumber,
          attachmentPath,
          fileName,
          message // !change this
        );
        console.log(`Image sent successfully to ${contact.name} (${contact.phone})`);
      } else {
        await client.sendFile(
          phoneNumber,
          attachmentPath,
          fileName,
          message
        );
        console.log(`File sent successfully to ${contact.name} (${contact.phone})`);
      }
    } catch (error) {
      console.error(`Error sending message/attachment to ${contact.name} (${contact.phone}):`, error);
    }

    await new Promise(resolve => setTimeout(resolve, 2000));
  }
}

venom
  .create('session-name', (base64Qr, asciiQR, attempts, urlCode) => {
    console.log('Number of attempts to read the qrcode: ', attempts);
    console.log('Terminal qrcode: ', asciiQR);
  }, undefined, options)
  .then((client) => start(client))
  .catch((erro) => {
    console.log(erro);
  });

async function start(client) {
  const excelFilePath = 'contacts.xlsx';
  const attachmentPath = 'attachment.jpg';  // Can be an image or any file
  const isImage = true;  // Set to false if sending a non-image file

  if (!fs.existsSync(excelFilePath)) {
    console.error('Excel file not found!');
    return;
  }
  if (!fs.existsSync(attachmentPath)) {
    console.error('Attachment file not found!');
    return;
  }

  const contacts = readAndCleanExcel(excelFilePath);
  
  console.log(`Loaded and cleaned ${contacts.length} contacts from Excel file.`);
  
  await sendMessagesWithAttachment(client, contacts, attachmentPath, isImage);
  
  console.log('All messages and attachments sent. You can close the script now.');
}