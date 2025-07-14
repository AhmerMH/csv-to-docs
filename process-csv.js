const fs = require('fs');
const docx = require('docx');
const { Document, Paragraph, Packer } = docx;

// Read the input file
const inputFile = 'training-file-20250713.txt';

// Generate timestamp for filename
const now = new Date();
const timestamp = now.getFullYear().toString() +
                 (now.getMonth() + 1).toString().padStart(2, '0') +
                 now.getDate().toString().padStart(2, '0') +
                 now.getHours().toString().padStart(2, '0') +
                 now.getMinutes().toString().padStart(2, '0');

const outputFile = `output_${timestamp}.docx`;

// Read the content of the file
fs.readFile(inputFile, 'utf8', (err, data) => {
    if (err) {
        console.error('Error reading file:', err);
        return;
    }

    // Process each line
    const lines = data.split('\n');
    const processedLines = lines.map((line, index) => {
        // Skip empty lines
        if (!line.trim()) return '';

        // Handle CSV parsing with consideration for quoted fields
        const processedLine = [];
        let field = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
                continue;
            }
            
            if (char === ',' && !inQuotes) {
                processedLine.push(field.trim());
                field = '';
                continue;
            }
            
            field += char;
        }
        
        // Add the last field
        processedLine.push(field.trim());

        // Filter out empty fields at the end
        while (processedLine.length > 0 && processedLine[processedLine.length - 1] === '') {
            processedLine.pop();
        }

        // For first row, just join with newlines
        if (index === 0) {
            return processedLine.join('\n');
        }
        // For other rows, add an extra newline at the end
        return processedLine.join('\n') + '\n';
    });

    // Join all processed lines with an extra newline between records (except after header)
    const output = processedLines.join('\n');

    // Create document
    const doc = new Document({
        sections: [{
            properties: {},
            children: output.split('\n').map(line => 
                new Paragraph({
                    text: line
                })
            )
        }]
    });

    // Save document
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFile(outputFile, buffer, (err) => {
            if (err) {
                console.error('Error writing file:', err);
                return;
            }
            console.log(`File has been processed successfully! Saved as ${outputFile}`);
        });
    });
}); 