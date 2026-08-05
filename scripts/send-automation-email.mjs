import fs from "node:fs";
import nodemailer from "nodemailer";

const smtpUsername = process.env.SMTP_USERNAME || "";
const smtpPassword = process.env.SMTP_PASSWORD || "";
const recipient = process.env.NOTIFICATION_EMAIL || "thearjunks@gmail.com";
const eventName = process.env.EMAIL_EVENT || "Automation";
const status = process.env.EMAIL_STATUS || "Status";
const details =
  process.env.EMAIL_DETAILS || "No additional details were provided.";

if (!smtpUsername || !smtpPassword) {
  throw new Error(
    "SMTP_USERNAME and SMTP_PASSWORD GitHub secrets are required for email notifications.",
  );
}

const attachmentPaths = (process.env.EMAIL_ATTACHMENTS || "")
  .split(",")
  .map((value) => value.trim())
  .filter((value) => value && fs.existsSync(value));

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: smtpUsername,
    pass: smtpPassword,
  },
});

const runUrl =
  process.env.GITHUB_SERVER_URL &&
  process.env.GITHUB_REPOSITORY &&
  process.env.GITHUB_RUN_ID
    ? `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`
    : "Not available";

await transporter.sendMail({
  from: `AK stc Labels Automation <${smtpUsername}>`,
  to: recipient,
  subject: `[AK stc Labels] ${eventName} - ${status}`,
  text: [
    `Event: ${eventName}`,
    `Status: ${status}`,
    `Time: ${new Date().toLocaleString("en-GB", { timeZone: "Asia/Kuwait" })} Kuwait time`,
    "",
    details,
    "",
    `GitHub Actions run: ${runUrl}`,
    "Production: https://labels.stcdigitalhub.com/labels",
  ].join("\n"),
  attachments: attachmentPaths.map((filePath) => ({
    filename: filePath.split(/[\\/]/).pop(),
    path: filePath,
  })),
});
