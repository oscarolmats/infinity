module.exports = (req, res) => {
  res.json({ 
    message: 'Test endpoint works!',
    timestamp: new Date().toISOString(),
    method: req.method,
    url: req.url
  });
};
