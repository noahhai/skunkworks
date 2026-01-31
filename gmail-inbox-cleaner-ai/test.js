require("dotenv").config({ path: ".env.local" });
const http = require("http");
const handler = require("./api/classify");

const TEST_PAYLOAD = {
  senders: [
    // Should DELETE — marketing/newsletters/notifications
    { email: "newsletter@marketing.com", name: "MarketCo", count: 50, subject: "50% off today!" },
    { email: "updates@linkedin.com", name: "LinkedIn", count: 120, subject: "You have 5 new notifications" },
    { email: "promo@retailstore.com", name: "RetailStore", count: 80, subject: "HUGE CLEARANCE SALE starts now" },
    { email: "digest@medium.com", name: "Medium Daily Digest", count: 200, subject: "Top stories for you" },
    // Should KEEP — personal/transactional/financial
    { email: "john.doe@gmail.com", name: "John Doe", count: 3, subject: "Hey, are we still on for lunch?" },
    { email: "noreply@amazon.com", name: "Amazon", count: 5, subject: "Your order has shipped" },
    { email: "billing@stripe.com", name: "Stripe", count: 2, subject: "Payment receipt for January" },
    { email: "security@bank.com", name: "Chase Bank", count: 1, subject: "Unusual sign-in activity detected" },
  ],
};

const EXPECTED_DELETE = new Set([
  "newsletter@marketing.com",
  "updates@linkedin.com",
  "promo@retailstore.com",
  "digest@medium.com",
]);

const EXPECTED_KEEP = new Set([
  "john.doe@gmail.com",
  "noreply@amazon.com",
  "billing@stripe.com",
  "security@bank.com",
]);

const server = http.createServer((req, res) => {
  if (req.method === "POST") {
    let data = "";
    req.on("data", (chunk) => (data += chunk));
    req.on("end", () => {
      try {
        req.body = JSON.parse(data);
      } catch (_) {
        req.body = {};
      }
      handler(req, res);
    });
  } else {
    handler(req, res);
  }
});

server.listen(0, async () => {
  const port = server.address().port;
  console.info(`Test server on port ${port}\n`);

  try {
    const res = await fetch(`http://localhost:${port}/api/classify`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(TEST_PAYLOAD),
    });

    const result = await res.json();
    console.info("Status:", res.status);
    console.info("Response:", JSON.stringify(result, null, 2));

    if (res.status !== 200) {
      console.info("\nFAIL: Expected status 200, got", res.status);
      process.exit(1);
    }

    const deleteSet = new Set((result.delete || []).map((e) => e.toLowerCase()));

    console.info("\n--- Results ---");
    let passed = true;

    for (const email of EXPECTED_DELETE) {
      const ok = deleteSet.has(email);
      console.info(ok ? "  PASS" : "  FAIL", `${email} — expected DELETE, got ${ok ? "DELETE" : "KEEP"}`);
      if (!ok) passed = false;
    }

    for (const email of EXPECTED_KEEP) {
      const ok = !deleteSet.has(email);
      console.info(ok ? "  PASS" : "  FAIL", `${email} — expected KEEP, got ${ok ? "KEEP" : "DELETE"}`);
      if (!ok) passed = false;
    }

    console.info(passed ? "\nAll checks passed." : "\nSome checks failed.");
    process.exit(passed ? 0 : 1);
  } catch (err) {
    console.info("Test error:", err.message);
    process.exit(1);
  } finally {
    server.close();
  }
});
