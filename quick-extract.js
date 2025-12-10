// Quick script to get the main calculator file
// Paste this in console on PropertyEngine

(async () => {
    const mainFile = 'https://propertyengine.co.uk/static/app/2912.b11dad1b44e7dbca900e.js';
    console.log('📥 Fetching main calculator file...');
    
    try {
        const response = await fetch(mainFile);
        const code = await response.text();
        
        console.log(`✅ Loaded ${(code.length / 1024).toFixed(2)} KB`);
        console.log('\n📋 Code saved to window.__PECode');
        console.log('💡 To copy: navigator.clipboard.writeText(window.__PECode)');
        
        window.__PECode = code;
        
        // Also try to find specific calculator functions
        console.log('\n🔍 Searching for calculator functions...');
        
        const patterns = [
            /function\s+(\w*[Cc]alc\w*|stampDuty|calculateBRR|calculateBTL)[\s(]/g,
            /const\s+(\w*[Cc]alc\w*|stampDuty|calculateBRR)[\s=]/g,
            /(\w*[Rr]ates?\w*)\s*[:=]\s*\{/g,
            /(\w*[Ss]tamp\w*)\s*[:=]/g
        ];
        
        const found = new Set();
        patterns.forEach(pattern => {
            let match;
            while ((match = pattern.exec(code)) !== null) {
                found.add(match[1]);
            }
        });
        
        console.log('Found potential functions/constants:', Array.from(found).slice(0, 20));
        
        return code;
    } catch (e) {
        console.error('Error:', e);
    }
})();

