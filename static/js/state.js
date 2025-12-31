// static/js/state.js - shared state

let listingCounter = 1;

export function showStatus(msg, type='info') {
    const color = type === 'error' ? 'text-danger' : 'text-success';
    $('#statusMsg').html(`<span class="${color}">${msg}</span>`);
    setTimeout(() => $('#statusMsg').empty(), 3000);
}

// Renumber listings
export async function renumberListings() {
    try {
        showStatus("⏳ Renumbering...");
        
        await Word.run(async (ctx) => {
            const paragraphs = ctx.document.body.paragraphs;
            ctx.load(paragraphs, 'text');
            await ctx.sync();
            
            const listingParagraphs = [];
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                if (paragraph.text.match(/Listing\s+\d+:/)) {
                    listingParagraphs.push(paragraph);
                }
            }
            
            for (let i = 0; i < listingParagraphs.length; i++) {
                const paragraph = listingParagraphs[i];
                const oldText = paragraph.text;
                const match = oldText.match(/Listing\s+\d+:(.*)/);
                const description = match ? match[1] : '';
                const newText = `Listing ${i + 1}:${description}`;
                paragraph.insertText(newText, 'Replace');
            }
            
            await ctx.sync();
            listingCounter = listingParagraphs.length + 1;
        });
        
        showStatus(`✅ Renumbered`);
    } catch (e) {
        console.error(e);
        showStatus("❌ Renumber failed: " + e.message, "error");
    }
}
