# ExcelJS

This is a fork of the "exceljs" package, which fixes the problem with writing a file using streams and not using RAM.
It solves the problem of writing large exel files

<a href="https://github.com/exceljs/exceljs">Original repo</a>

Guys from ExcelJS promise this fix will be released with next major version (v5) as it has some breaking changes

<a href="https://github.com/exceljs/exceljs/pull/2558">Original PR with applied fix</a>

# Installation

```bash
npm install @dima-supy-io/exceljs
```

# Whats new!
To use streams correctly just write:

```javascript
import * as fs from 'fs';
import { stream } from '@dima-supy-io/exceljs';

const output_file_name = "/test.xlsx";

const writeStream = fs.createWriteStream(output_file_name, { flags: 'w' });
const wb = new stream.xlsx.WorkbookWriter({ stream: writeStream });
const worksheet = wb.addWorksheet("test");

const headers = Array.from({length: 256}, (_, i) => i + 1).map((i) => 'test' + i);

for (let i = 0; i < 100000; i++) {
  const row = headers.map((header) => header + '|' + i);
  await worksheet.addRow(row).commit(); // This row will be immediately written to disk and will not clog RAM.
}

await worksheet.commit(); // This is not necessary because await wb.commit() is used, but you can also write to disk not row by row, but worksheet by worksheet.
await wb.commit();
```


or with Google Storage Bucket:

```javascript
import { PassThrough } from 'stream';
import { stream } from '@dima-supy-io/exceljs';
import { Storage } from '@google-cloud/storage';
import WorkbookWriter = stream.xlsx.WorkbookWriter;

const writeStream = new PassThrough();
const options = {
    stream: writeStream,
    useStyles: true,
    useSharedStrings: true,
};
const workbook = WorkbookWriter(options);

const uploadStream = new Storage().bucket('gcp-bucket').file('upload-filename').createWriteStream({
    metadata: {
        metadata,
    },
});

writeStream.pipe(uploadStream);

const worksheet = workbook.addWorksheet("test");

const headers = Array.from({length: 256}, (_, i) => i + 1).map((i) => 'test' + i);

for (let i = 0; i < 100000; i++) {
  const row = headers.map((header) => header + '|' + i);
  await worksheet.addRow(row).commit(); // This row will be immediately written to GCP Storage and will not clog RAM.
}

await worksheet.commit(); // This is not necessary because await wb.commit() is used, but you can also dispatch to GCP Storage Bucket not row by row, but worksheet by worksheet.
await wb.commit();

await new Promise((resolve, reject) => {
    uploadStream.on('finish', resolve);
    uploadStream.on('error', reject);
});
```
