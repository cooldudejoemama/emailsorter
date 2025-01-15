require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { OpenAI } = require('openai');

const app = express();
app.use(cors());
app.use(express.json());

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

app.post('/classify', async (req, res) => {
    try {
        const { subject, body } = req.body;
        
        const completion = await openai.chat.completions.create({
            model: "gpt-3.5-turbo",
            messages: [
                {
                    role: "system",
                    content: "You are an email classifier. Respond with exactly one of these categories: Personal, Work, Finance, Shopping, Social, Other"
                },
                {
                    role: "user",
                    content: `Classify this email:\nSubject: ${subject}\nBody: ${body}`
                }
            ],
            temperature: 0.3,
            max_tokens: 50
        });

        const folder = completion.choices[0].message.content.trim();
        
        res.json({ folder });
    } catch (error) {
        console.error('Classification error:', error);
        res.status(500).json({ error: 'Classification failed' });
    }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});