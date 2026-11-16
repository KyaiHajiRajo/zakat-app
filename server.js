const app = require("./app");

// Get port from environment variables or use default
const PORT = process.env.PORT || 3000;

// Start server
const server = app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log(`📅 Started at: ${new Date().toLocaleString("id-ID")}`);
  console.log(`🌍 Environment: ${process.env.NODE_ENV || "development"}`);
});

// Log aktivitas lebih dari 1 tahun dibersihkan saat server menyala
app.locals.db.execute("DELETE FROM log_aktivitas WHERE created_at < (NOW() - INTERVAL 365 DAY)").catch(() => {});

// Graceful shutdown
function shutdown(signal) {
  console.log(`📥 ${signal} received. Shutting down gracefully...`);
  server.close(() => {
    console.log("✅ Process terminated");
    process.exit(0);
  });
  server.closeIdleConnections();
  setTimeout(() => process.exit(0), 3000).unref();
}

process.on("SIGTERM", () => shutdown("SIGTERM"));
process.on("SIGINT", () => shutdown("SIGINT"));

module.exports = server;
