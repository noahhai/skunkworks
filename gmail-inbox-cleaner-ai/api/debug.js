module.exports = function handler(req, res) {
  res.setHeader("Content-Type", "application/json");
  res.statusCode = 200;
  res.end(JSON.stringify({
    method: req.method,
    bodyType: typeof req.body,
    body: req.body,
    headers: req.headers,
  }));
};
