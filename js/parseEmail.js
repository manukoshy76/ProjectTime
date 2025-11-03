async function parseEmail(file) {
  const fileName = file.name.replace(/\.msg$/i, '').replace(/\.eml$/i, '');
  const email = {
    id: Date.now(),
    subject: fileName,
    from: "Unknown Sender",
    body: "",
    date: new Date().toISOString().split("T")[0],
    read: false,
    notes: "",
    fileName: file.name,
  };

  const ext = file.name.split('.').pop().toLowerCase();

  // ✅ MSG Reader handling
  if (ext === "msg") {
    try {
      if (!window.MSGReader) {
        console.error("MSGReader not loaded!");
        alert("MSGReader not loaded — check script order");
        return email;
      }

      const buffer = await file.arrayBuffer();
      const msgReader = new MSGReader(buffer);
      const data = msgReader.getFileData();

      if (data.error) throw new Error("Invalid or unreadable .msg file");

      email.from = data.senderName
        ? `${data.senderName} <${data.senderEmail}>`
        : data.senderEmail || email.from;

      email.subject = data.subject || fileName;

      // Try to read date
      if (data.headers) {
        const match = data.headers.match(/Date:\s*(.*)/i);
        if (match && match[1]) {
          const parsed = new Date(match[1]);
          if (!isNaN(parsed.getTime()))
            email.date = parsed.toISOString().split("T")[0];
        }
      }

      email.body = data.body || data.bodyHTML || "(no content)";
      console.log("✅ MSG parsed successfully:", data);
      return email;
    } catch (err) {
      console.error("MSG parse failed:", err);
    }
  }

  // ✅ fallback for .eml or text files
  try {
    const text = await file.text();
    const from = text.match(/^From:\s*(.*)$/im);
    const subject = text.match(/^Subject:\s*(.*)$/im);
    const date = text.match(/^Date:\s*(.*)$/im);
    if (from) email.from = from[1].trim();
    if (subject) email.subject = subject[1].trim();
    if (date) {
      const d = new Date(date[1]);
      if (!isNaN(d.getTime())) email.date = d.toISOString().split("T")[0];
    }
    email.body = text.split(/\r?\n\r?\n/).slice(1).join("\n").slice(0, 4000);
  } catch (err) {
    console.error("Fallback parse failed:", err);
  }

  return email;
}
