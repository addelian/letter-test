import React, { useState } from 'react';
import { utils, read } from 'xlsx';
import mustache from 'mustache';
import './LetterGenerator.css';

function LetterGenerator({ template }) {

    const [lettersGenerated, setLettersGenerated] = useState(false);
    const [error, setError] = useState(null);
    const [file, setFile] = useState(null);

    function handleFileSelect(event) {
        const inputFile = event.target.files[0];
        setFile(inputFile);
    }

    function generateLetters() {
        try {
            const reader = new FileReader();
            reader.readAsArrayBuffer(file);
            reader.onload = (event) => {
                const data = new Uint8Array(event.target.result);
                const workbook = read(data, { type: "array" });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const recipients = utils.sheet_to_json(worksheet, { raw: false });
                const sendEmail = (person) => {
                    const letter = mustache.render(template, person);
                    const recipientEmail = person.email;
                    const subject = 'A message from Nic and ChatGPT';
                    const body = letter;
                    window.open(`mailto:${recipientEmail}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`);
                }
                for (let i = 0; i < recipients.length; i++) {
                    setTimeout(() => {
                        sendEmail(recipients[i]);
                    }, i * 2000)
                }
            };
            setLettersGenerated(true);
        } catch (error) {
            setError(error);
        }
    }

    return (
        <div className="letter-generator-container">
            <input type="file" accept=".xlsx" onChange={(e) => handleFileSelect(e)} />
            {!lettersGenerated && !error && <button onClick={generateLetters}>Generate letters</button>}
            {lettersGenerated && <p>Letters generated successfully! Visit each new browser tab to send your emails.</p>}
            {error && <p>Error: {error.message}</p>}
        </div>
    );
}

export default LetterGenerator;