const express = require('express');
const multer = require('multer');
const docx = require('docx');
const { Document, Paragraph, Packer } = docx;

const app = express();
// Configure multer to use memory storage instead of disk
const upload = multer({ storage: multer.memoryStorage() });

// Serve static files from public directory
app.use(express.static('public'));

app.post('/convert', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).send('No file uploaded');
        }

        // Read the uploaded file from buffer
        const fileContent = req.file.buffer.toString('utf8');
        
        // Process the CSV content
        const lines = fileContent.split('\n');
        const processedLines = lines.map((line, index) => {
            if (!line.trim()) return '';

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

        // Join all processed lines
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

        // Generate document buffer
        const buffer = await Packer.toBuffer(doc);

        // Send the document
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buffer);

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error processing file');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
}); 