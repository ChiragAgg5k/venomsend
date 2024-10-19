const venom = require('venom-bot');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// Configuration options
const config = {
  excelFilePath: 'contacts.xlsx',
  attachmentPath: 'attachment.jpg',
  message: `Hello {name},

This is a multi-line message from Venom Bot!

We hope you're having a great day.

Best regards,
The Venom Bot Team`,
  delayBetweenMessages: 2000,  // Delay in milliseconds
};

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
  return config.message.replace('{name}', contact.name);
}

async function sendMessagesWithAttachment(client, contacts, sendOption, textAttachmentOption) {
  for (const contact of contacts) {
    const phoneNumber = `91${contact.phone}@c.us`;
    const message = constructMessage(contact);
    
    try {
      switch (sendOption) {
        case '1':
          await client.sendText(phoneNumber, message);
          console.log(`Text message sent successfully to ${contact.name} (${contact.phone})`);
          break;
        case '2':
          await client.sendFile(
            phoneNumber,
            config.attachmentPath,
            path.basename(config.attachmentPath),
            ""
          );
          console.log(`File sent successfully to ${contact.name} (${contact.phone})`);
          break;
        case '3':
          if (textAttachmentOption === '1') {
            await client.sendFile(
              phoneNumber,
              config.attachmentPath,
              path.basename(config.attachmentPath),
              message
            );
            console.log(`Text and file sent together successfully to ${contact.name} (${contact.phone})`);
          } else {
            await client.sendText(phoneNumber, message);
            console.log(`Text message sent successfully to ${contact.name} (${contact.phone})`);
            await new Promise(resolve => setTimeout(resolve, 1000)); // Small delay between text and attachment
            await client.sendFile(
              phoneNumber,
              config.attachmentPath,
              path.basename(config.attachmentPath),
              ""
            );
            console.log(`File sent separately and successfully to ${contact.name} (${contact.phone})`);
          }
          break;
      }
    } catch (error) {
      console.error(`Error sending message/attachment to ${contact.name} (${contact.phone}):`, error);
    }

    await new Promise(resolve => setTimeout(resolve, config.delayBetweenMessages));
  }
}

function askUser(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  return new Promise(resolve => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer);
    });
  });
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
  if (!fs.existsSync(config.excelFilePath)) {
    console.error('Excel file not found!');
    return;
  }
  if (!fs.existsSync(config.attachmentPath)) {
    console.error('Attachment file not found!');
    return;
  }

  const contacts = readAndCleanExcel(config.excelFilePath);
  
  console.log(`Loaded and cleaned ${contacts.length} contacts from Excel file.`);

  const sendOption = await askUser(
    "Choose an option:\n1. Send only text\n2. Send only image/attachment\n3. Send text with image\nEnter your choice (1, 2, or 3): "
  );

  if (!['1', '2', '3'].includes(sendOption)) {
    console.error('Invalid option selected. Exiting...');
    return;
  }

  let textAttachmentOption = '1';
  if (sendOption === '3') {
    textAttachmentOption = await askUser(
      "How do you want to send text and attachment?\n1. As a single message\n2. Text first, then attachment\nEnter your choice (1 or 2): "
    );
    if (!['1', '2'].includes(textAttachmentOption)) {
      console.error('Invalid option selected. Exiting...');
      return;
    }
  }
  
  await sendMessagesWithAttachment(client, contacts, sendOption, textAttachmentOption);
  
  console.log('All messages and attachments sent. You can close the script now.');
}