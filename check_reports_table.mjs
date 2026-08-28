import axios from 'axios';

async function run() {
    try {
        console.log("Calling local API on port 9001...");
        const res = await axios.get("http://localhost:9001/api/reports/externalAPI/godownSummaryInstock", {
            params: {
                Fromdate: "2026-08-01",
                Todate: "2026-08-28"
            }
        });
        
        console.log("API Response Success:", res.data?.success);
        const data = res.data?.data;
        if (Array.isArray(data)) {
            console.log("Sample row (array):", data[0]);
        } else if (data && typeof data === 'object') {
            const list = data.Data4 || data.data || [];
            console.log("Sample row (object Data4):", list[0]);
        } else {
            console.log("Raw response:", res.data);
        }
    } catch (e) {
        console.error("API Call error:", e.message);
    }
}

run();
