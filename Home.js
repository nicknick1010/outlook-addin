function submitApiKey() {
    const apiKey = document.getElementById('apiKey').value;
    localStorage.setItem('openai_api_key', apiKey);
    document.getElementById('status').innerText = 'API Key saved. You can now send emails.';
}

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, onItemSend);
    }
});

async function analyzeEmailContent(content) {
    let apiKey = localStorage.getItem('openai_api_key');
    if (!apiKey) {
        alert('Please set your OpenAI API key in the add-in pane.');
        return '';
    }

    const response = await fetch('https://api.openai.com/v1/engines/davinci-codex/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            prompt: `Analyze this email content: ${content}`,
            max_tokens: 150
        })
    });

    const data = await response.json();
    return data.choices[0].text;
}

async function onItemSend(eventArgs) {
    const item = Office.context.mailbox.item;
    const body = await item.body.getAsync(Office.CoercionType.Text);
    const analysis = await analyzeEmailContent(body.value);

    if (analysis.includes('potential issue')) {
        eventArgs.completed({ allowEvent: false });
        Office.context.mailbox.item.notificationMessages.addAsync('alert', {
            type: 'errorMessage',
            message: 'Potential issue detected with the email content. Please review before sending.'
        });
    } else {
        eventArgs.completed({ allowEvent: true });
    }
}
