import Imap from "imap";
import fs from "fs/promises";
import xlsx from "xlsx";
import {PdfReader} from 'pdfreader';
import {simpleParser} from 'mailparser'
import {TextDecoder} from 'util'
const imapConfig = {
  user: "aravindpj4554@gmail.com",
  password: "dvypoenqrukoopao",
  host: "imap.gmail.com",
  port: 993,
  tls: true,
  tlsOptions: {
    rejectUnauthorized: false, // Allow self-signed certificates
  },
};
const getMail = function () {
  console.log("--email checking--");
  const imap = new Imap(imapConfig);
  imap.once("ready", () => {
    imap.openBox("INBOX", false, () => {
      imap.search(["UNSEEN", ["SINCE", new Date()]], (err, results) => {
        const f = imap.fetch(results, {
          bodies: "",
          struct: true,
        });
        f.on("message", (msg) => {
          msg.on("body", (stream) => {
            simpleParser(stream, async (err, parsed) => {
              const { from, to, subject, text, attachments } = parsed;
              // console.log(attachments)
              if (attachments.length > 0) {
                console.log('-working-attachments-')
                for (const attachment of attachments) {
                  let fileType=attachment.filename.split('.').pop()

                  switch(fileType){
                    case 'xlsx':
                      console.log(`Processing XLSX file: ${attachment.filename}`);
                      const workbook = xlsx.read(attachment.content);
                      // Now, you can work with the Excel workbook
                      const sheetName = workbook.SheetNames[0];
                      const sheet = workbook.Sheets[sheetName];
                      let data = xlsx.utils.sheet_to_json(sheet);
                      console.log(JSON.stringify(data));
                      break;
                    case 'pdf':
                      console.log(`Processing PDF file: ${attachment.filename}`);
                      new PdfReader().parseBuffer(attachment.content, (err, item) => {
                        if (err) console.error("error:", err);
                        else if (!item) console.warn("end of buffer");
                        else if (item.text) console.log(item.text)
                      });
                      break;
                    case 'doc':
                      console.log(`Processing DOC file: ${attachment.filename}`);
                      const decoder = new TextDecoder('utf-8');
                      const text = decoder.decode(attachment.content);
                      // console.log(`MSDoc attachment converted to text: ${text}`);             
                    }
                }
              } else {
                console.log('-working-messageText-')
                const pattern = /\* ([^\n=]+?)\s*=\s*([\d\w\s]+)(?=\n|\s*$)/g;

                const productsData = new Set();
                let match;
                while ((match = pattern.exec(text)) !== null) {
                  let product = match[1].trim();
                  let quantity = match[2].trim();
                  productsData.add(product);
                }
                console.log(productsData);
              }
            });
          });
        });
        f.once("error", (ex) => {
          return Promise.reject(ex);
        });
        f.once("end", () => {
          console.log("Done fetching all messages!");
          imap.end();
        });
      });
    });
  });

  imap.once("error", (err) => {
    console.log(err);
  });

  imap.once("end", () => {
    console.log("Connection ended");
  });

  imap.connect();
};

getMail();
