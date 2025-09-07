import { Lead } from '../types';
import { NAME_ALIASES, EMAIL_ALIASES, MOBILE_ALIASES } from '../constants';

// Since XLSX is loaded from a CDN, we declare it to satisfy TypeScript
declare const XLSX: any;

const findKeyByAliases = (obj: { [key: string]: any }, aliases: string[]): string | null => {
  const lowerCaseKeys = Object.keys(obj).map(k => k.toLowerCase().trim());
  for (const alias of aliases) {
    const keyIndex = lowerCaseKeys.indexOf(alias);
    if (keyIndex !== -1) {
      return Object.keys(obj)[keyIndex];
    }
  }
  return null;
};


export const processExcelFiles = (files: FileList): Promise<Lead[]> => {
  return new Promise(async (resolve, reject) => {
    const allLeads: Lead[] = [];
    const filePromises = Array.from(files).map(file => {
      return new Promise<Lead[]>((resolveFile, rejectFile) => {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'array' });
            const fileLeads: Lead[] = [];

            workbook.SheetNames.forEach((sheetName: string) => {
              const worksheet = workbook.Sheets[sheetName];
              const json: any[] = XLSX.utils.sheet_to_json(worksheet);

              if (json.length > 0) {
                  json.forEach(row => {
                    const nameKey = findKeyByAliases(row, NAME_ALIASES);
                    const emailKey = findKeyByAliases(row, EMAIL_ALIASES);
                    const mobileKey = findKeyByAliases(row, MOBILE_ALIASES);

                    // A row is only valid if it has a mobile number that can be formatted to 10 digits.
                    if (mobileKey && row[mobileKey]) {
                        const rawMobile = String(row[mobileKey]).trim();
                        // Remove all non-digit characters to handle formats like (123) 456-7890
                        const cleanedMobile = rawMobile.replace(/\D/g, ''); 
                        let finalMobile = '';

                        if (cleanedMobile.length > 10) {
                            // If more than 10 digits, take the last 10.
                            finalMobile = cleanedMobile.slice(-10);
                        } else if (cleanedMobile.length === 10) {
                            // If exactly 10 digits, it's valid.
                            finalMobile = cleanedMobile;
                        }

                        // If finalMobile is set, we have a valid 10-digit number.
                        // Rows with mobile numbers < 10 digits are automatically skipped.
                        if (finalMobile) {
                             const lead: Lead = {
                                name: nameKey ? String(row[nameKey]).trim() : '',
                                email: emailKey ? String(row[emailKey]).trim() : '',
                                mobile: finalMobile,
                            };
                            fileLeads.push(lead);
                        }
                    }
                });
              }
            });
            resolveFile(fileLeads);
          } catch (error) {
            console.error("Error processing file:", file.name, error);
            rejectFile(`Error processing file: ${file.name}`);
          }
        };
        reader.onerror = (error) => rejectFile(error);
        reader.readAsArrayBuffer(file);
      });
    });

    try {
      const results = await Promise.all(filePromises);
      results.forEach(leads => allLeads.push(...leads));
      resolve(allLeads);
    } catch (error) {
      reject(error);
    }
  });
};

export const exportToExcel = (leads: Lead[], fileName: string): void => {
  const dataToExport = leads.map(lead => ({
    'Full Name': lead.name,
    'Email Address': lead.email,
    'Mobile Number': lead.mobile
  }));

  const worksheet = XLSX.utils.json_to_sheet(dataToExport);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Consolidated Leads');
  
  const max_widths = dataToExport.reduce((widths, row) => ({
      name: Math.max(widths.name, row['Full Name']?.length || 10),
      email: Math.max(widths.email, row['Email Address']?.length || 15),
      mobile: Math.max(widths.mobile, row['Mobile Number']?.length || 15),
  }), {name: 15, email: 20, mobile: 15});
  
  worksheet["!cols"] = [ { wch: max_widths.name }, { wch: max_widths.email }, { wch: max_widths.mobile } ];

  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};