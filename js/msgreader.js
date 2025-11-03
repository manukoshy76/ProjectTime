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

  if (ext === "msg") {
    try {
      if (!window.MSGReader) {
        console.error("MSGReader not loaded!");
        return email;
      }

      const arrayBuffer = await file.arrayBuffer();
      const msgReader = new MSGReader(arrayBuffer);
      const data = msgReader.getFileData();

      if (data.error) throw new Error("Invalid .msg file");

      email.from = data.senderName
        ? `${data.senderName} <${data.senderEmail}>`
        : (data.senderEmail || "Unknown Sender");

      email.subject = (data.subject || fileName)
        .replace(/^(RE|FW|FWD):\s*/i, "")
        .trim();

      if (data.headers) {
        const match = data.headers.match(/Date:\s*(.*)/i);
        if (match) {
          const parsedDate = new Date(match[1]);
          if (!isNaN(parsedDate.getTime()))
            email.date = parsedDate.toISOString().split("T")[0];
        }
      } else if (data.messageDeliveryTime) {
        const d = new Date(data.messageDeliveryTime);
        if (!isNaN(d.getTime())) email.date = d.toISOString().split("T")[0];
      }

      email.body = data.body || data.bodyHTML || "(no content)";

      return email;
    } catch (err) {
      console.error("MSG parse failed:", err);
      return email;
    }
  }

  // fallback for .eml/text
  try {
    const text = await file.text();
    email.from = (text.match(/^From:\s*(.*)$/im) || [])[1] || email.from;
    email.subject = (text.match(/^Subject:\s*(.*)$/im) || [])[1] || email.subject;
    const dateStr = (text.match(/^Date:\s*(.*)$/im) || [])[1];
    if (dateStr) {
      const d = new Date(dateStr);
      if (!isNaN(d.getTime())) email.date = d.toISOString().split("T")[0];
    }
    email.body = text.split(/\r?\n\r?\n/).slice(1).join("\n").slice(0, 4000);
  } catch (err) {
    console.error("Fallback parse failed:", err);
  }

  return email;
}
