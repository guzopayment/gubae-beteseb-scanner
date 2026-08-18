import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { Html5Qrcode } from "html5-qrcode";
import axios from "axios";
import ExcelJS from "exceljs";
import {
  CheckCircle2,
  Download,
  LogIn,
  LogOut,
  RefreshCw,
  ScanLine,
  Users,
  XCircle,
} from "lucide-react";
import "./style.css";

const API =
  import.meta.env.VITE_API_BASE_URL ||
  import.meta.env.VITE_API_URL ||
  "http://localhost:10000/api";
const api = axios.create({ baseURL: API });
api.interceptors.request.use((config) => {
  const token = localStorage.getItem("adminToken");
  if (token) config.headers.Authorization = `Bearer ${token}`;
  return config;
});

function Login({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const r = await api.post("/auth/login", { email, password });
      localStorage.setItem("adminToken", r.data.token);
      onLogin();
    } catch (err) {
      setError(err.response?.data?.message || "Login failed");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="center">
      <form className="card login" onSubmit={submit}>
        <h1>Gubae Attendance</h1>
        <p>Authorized event scanner</p>
        <input
          type="email"
          placeholder="Admin email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          required
        />
        <input
          type="password"
          placeholder="Password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          required
        />
        <button disabled={loading}>
          <LogIn size={18} />
          {loading ? "Signing in…" : "Sign in"}
        </button>
        {error && <div className="error">{error}</div>}
      </form>
    </div>
  );
}

function Scanner() {
  const scannerRef = useRef(null);
  const runningRef = useRef(false);
  const processingRef = useRef(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [summary, setSummary] = useState(null);
  const [scanning, setScanning] = useState(false);
  const [orgPage, setOrgPage] = useState(1);
  const [resetOpen, setResetOpen] = useState(false);
  const [resetPhrase, setResetPhrase] = useState("");
  const [resetCount, setResetCount] = useState("");
  const [resetBusy, setResetBusy] = useState(false);
  const audioContextRef = useRef(null);

  const vibrateRejected = () => {
    try {
      if (
        typeof navigator !== "undefined" &&
        typeof navigator.vibrate === "function"
      ) {
        navigator.vibrate([180, 90, 180]);
      }
    } catch (vibrationError) {
      console.warn("Scanner vibration unavailable", vibrationError);
    }
  };

  const playScanSound = (accepted = true) => {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = audioContextRef.current || new AudioCtx();
      audioContextRef.current = ctx;
      if (ctx.state === "suspended") ctx.resume();

      const oscillator = ctx.createOscillator();
      const gain = ctx.createGain();
      oscillator.connect(gain);
      gain.connect(ctx.destination);

      const now = ctx.currentTime;
      if (accepted) {
        oscillator.frequency.setValueAtTime(880, now);
        oscillator.frequency.setValueAtTime(1320, now + 0.09);
        gain.gain.setValueAtTime(0.0001, now);
        gain.gain.exponentialRampToValueAtTime(0.28, now + 0.015);
        gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.22);
        oscillator.start(now);
        oscillator.stop(now + 0.24);
      } else {
        oscillator.frequency.setValueAtTime(240, now);
        oscillator.frequency.setValueAtTime(170, now + 0.12);
        gain.gain.setValueAtTime(0.0001, now);
        gain.gain.exponentialRampToValueAtTime(0.32, now + 0.015);
        gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.3);
        oscillator.start(now);
        oscillator.stop(now + 0.32);
      }
    } catch (soundError) {
      console.warn("Scanner sound unavailable", soundError);
    }
  };

  const stop = async () => {
    if (scannerRef.current && runningRef.current) {
      try {
        await scannerRef.current.stop();
      } catch {}
      try {
        scannerRef.current.clear();
      } catch {}
    }
    runningRef.current = false;
    setScanning(false);
  };

  const resetAttendance = async () => {
    if (resetPhrase !== "RESET ATTENDANCE") {
      setError('For safety, type exactly "RESET ATTENDANCE".');
      return;
    }
    if (String(resetCount) !== String(summary?.totalPresent ?? "")) {
      setError(
        `For safety, enter the current Present count exactly: ${summary?.totalPresent ?? 0}`,
      );
      return;
    }

    setResetBusy(true);
    setError("");
    try {
      const r = await api.post("/bookings/attendance/reset", {
        confirmation: resetPhrase,
        expectedPresent: Number(resetCount),
      });
      setResult(null);
      setResetOpen(false);
      setResetPhrase("");
      setResetCount("");
      setError("");
      await loadSummary();
      alert(r.data?.message || "Attendance reset successfully.");
    } catch (err) {
      if (err.response?.status === 401) logout();
      else
        setError(err.response?.data?.message || "Unable to reset attendance");
    } finally {
      setResetBusy(false);
    }
  };

  const downloadAttendanceList = async () => {
    setError("");
    try {
      const r = await api.get("/bookings/attendance/list");
      const rows = r.data?.participants || [];
      const generatedAt = new Date(r.data?.generatedAt || Date.now());
      const total = r.data?.total ?? rows.length;
      const present =
        r.data?.present ??
        rows.filter((row) => row.status === "Present").length;
      const absent = r.data?.absent ?? total - present;

      const organizationMap = new Map();
      rows.forEach((row) => {
        const organization =
          String(row.organization || "Unknown").trim() || "Unknown";
        if (!organizationMap.has(organization)) {
          organizationMap.set(organization, {
            organization,
            totalRegistered: 0,
            totalPresent: 0,
            totalAbsent: 0,
            men: 0,
            women: 0,
          });
        }
        const item = organizationMap.get(organization);
        item.totalRegistered += 1;
        if (String(row.status).toLowerCase() === "present")
          item.totalPresent += 1;
        else item.totalAbsent += 1;
        const sex = String(row.sex || "")
          .trim()
          .toLowerCase();
        if (sex === "male" || sex === "ወንድ") item.men += 1;
        if (sex === "female" || sex === "ሴት") item.women += 1;
      });
      const organizationRows = Array.from(organizationMap.values());

      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Gubae Attendance System";
      workbook.lastModifiedBy = "Gubae Attendance System";
      workbook.created = new Date();
      workbook.modified = new Date();
      workbook.properties.subject = "Participant Attendance Report";
      workbook.properties.title = "Gubae Participant Attendance Report";
      workbook.properties.company = "Gubae";

      const sheet = workbook.addWorksheet("Attendance Report", {
        views: [{ state: "frozen", ySplit: 8 }],
        pageSetup: {
          orientation: "landscape",
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          paperSize: 9,
        },
        properties: { defaultRowHeight: 20 },
      });

      sheet.mergeCells("A1:G1");
      const title = sheet.getCell("A1");
      title.value = "GUBAE PARTICIPANT ATTENDANCE REPORT";
      title.font = {
        name: "Aptos Display",
        size: 18,
        bold: true,
        color: { argb: "FFFFFFFF" },
      };
      title.alignment = { vertical: "middle", horizontal: "center" };
      title.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "003B46" },
      };
      sheet.getRow(1).height = 34;

      sheet.mergeCells("A2:G2");
      const subtitle = sheet.getCell("A2");
      subtitle.value = `Generated: ${generatedAt.toLocaleString()}`;
      subtitle.font = {
        name: "Aptos",
        size: 10,
        italic: true,
        color: { argb: "555555" },
      };
      subtitle.alignment = { horizontal: "center", vertical: "middle" };
      sheet.getRow(2).height = 24;

      const metricLabels = [
        "Total Registered",
        "Present",
        "Absent",
        "Attendance Rate",
      ];
      const metricValues = [
        total,
        present,
        absent,
        total ? present / total : 0,
      ];
      const metricRanges = ["A4:B5", "C4:D5", "E4:F5", "G4:G5"];
      metricRanges.forEach((range, i) => sheet.mergeCells(range));
      const metricCells = ["A4", "C4", "E4", "G4"];
      const metricFills = ["EAF4F4", "E8F5E9", "FDECEC", "FFF4D6"];
      metricCells.forEach((addr, i) => {
        const cell = sheet.getCell(addr);
        cell.value = `${metricLabels[i]}\n${i === 3 ? `${(metricValues[i] * 100).toFixed(1)}%` : metricValues[i]}`;
        cell.font = {
          name: "Aptos",
          size: 12,
          bold: true,
          color: { argb: "003B46" },
        };
        cell.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: metricFills[i] },
        };
        cell.border = {
          top: { style: "thin", color: { argb: "D5DDE0" } },
          bottom: { style: "thin", color: { argb: "D5DDE0" } },
          left: { style: "thin", color: { argb: "D5DDE0" } },
          right: { style: "thin", color: { argb: "D5DDE0" } },
        };
      });
      sheet.getRow(4).height = 24;
      sheet.getRow(5).height = 24;

      const headerRow = 7;
      const headers = [
        "No.",
        "Name",
        "Organization",
        "Phone",
        "Sex",
        "Status",
        "Checked In At",
      ];
      const header = sheet.getRow(headerRow);
      header.values = headers;
      header.height = 28;
      header.eachCell((cell) => {
        cell.font = {
          name: "Aptos",
          size: 11,
          bold: true,
          color: { argb: "FFFFFFFF" },
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "006D77" },
        };
        cell.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
        cell.border = {
          bottom: { style: "medium", color: { argb: "003B46" } },
        };
      });

      rows.forEach((row, index) => {
        const excelRow = sheet.addRow([
          row.number ?? index + 1,
          row.name ?? "",
          row.organization ?? "",
          row.phone ?? "",
          row.sex ?? "",
          row.status ?? "Absent",
          row.checkedInAt ? new Date(row.checkedInAt) : "",
        ]);
        excelRow.height = 22;
        excelRow.eachCell((cell) => {
          cell.font = { name: "Aptos", size: 10, color: { argb: "222222" } };
          cell.alignment = { vertical: "middle", wrapText: true };
          cell.border = {
            bottom: { style: "hair", color: { argb: "D9E2E5" } },
          };
        });
        if (index % 2 === 1) {
          excelRow.eachCell((cell) => {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "F4F8F8" },
            };
          });
        }
        const statusCell = excelRow.getCell(6);
        const isPresent = String(row.status || "").toLowerCase() === "present";
        statusCell.font = {
          name: "Aptos",
          size: 10,
          bold: true,
          color: { argb: isPresent ? "1B5E20" : "9E2A2B" },
        };
        statusCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: isPresent ? "E8F5E9" : "FDECEC" },
        };
        statusCell.alignment = { horizontal: "center", vertical: "middle" };
        const dateCell = excelRow.getCell(7);
        if (row.checkedInAt) dateCell.numFmt = "dd-mmm-yyyy hh:mm AM/PM";
      });

      sheet.autoFilter = {
        from: `A${headerRow}`,
        to: `G${headerRow + rows.length}`,
      };
      sheet.getColumn(1).width = 8;
      sheet.getColumn(2).width = 28;
      sheet.getColumn(3).width = 36;
      sheet.getColumn(4).width = 18;
      sheet.getColumn(5).width = 14;
      sheet.getColumn(6).width = 16;
      sheet.getColumn(7).width = 24;

      sheet.getColumn(1).alignment = { horizontal: "center" };
      sheet.getColumn(4).alignment = { horizontal: "center" };
      sheet.getColumn(5).alignment = { horizontal: "center" };
      sheet.getColumn(6).alignment = { horizontal: "center" };

      const orgSheet = workbook.addWorksheet("Organization Summary", {
        views: [{ state: "frozen", ySplit: 4 }],
        pageSetup: {
          orientation: "landscape",
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          paperSize: 9,
        },
      });
      orgSheet.mergeCells("A1:F1");
      orgSheet.getCell("A1").value = "ATTENDANCE BY ORGANIZATION";
      orgSheet.getCell("A1").font = {
        name: "Aptos Display",
        size: 16,
        bold: true,
        color: { argb: "FFFFFFFF" },
      };
      orgSheet.getCell("A1").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      orgSheet.getCell("A1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "003B46" },
      };
      orgSheet.getRow(1).height = 32;
      orgSheet.mergeCells("A2:F2");
      orgSheet.getCell("A2").value =
        `Present: ${present} | Absent: ${absent} | Total Registered: ${total}`;
      orgSheet.getCell("A2").alignment = { horizontal: "center" };
      orgSheet.getCell("A2").font = {
        name: "Aptos",
        size: 10,
        color: { argb: "555555" },
      };
      orgSheet.getRow(4).values = [
        "Organization",
        "Registered",
        "Present",
        "Absent",
        "Men",
        "Women",
      ];
      orgSheet.getRow(4).eachCell((cell) => {
        cell.font = {
          name: "Aptos",
          size: 11,
          bold: true,
          color: { argb: "FFFFFFFF" },
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "006D77" },
        };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });
      organizationRows.forEach((item, index) => {
        const rr = orgSheet.addRow([
          item.organization,
          item.totalRegistered,
          item.totalPresent,
          item.totalAbsent,
          item.men,
          item.women,
        ]);
        rr.eachCell((cell) => {
          cell.font = { name: "Aptos", size: 10 };
          cell.alignment = { vertical: "middle" };
        });
        if (index % 2 === 1)
          rr.eachCell((cell) => {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "F4F8F8" },
            };
          });
      });
      orgSheet.autoFilter = {
        from: "A4",
        to: `F${4 + organizationRows.length}`,
      };
      orgSheet.getColumn(1).width = 42;
      orgSheet.getColumn(2).width = 14;
      orgSheet.getColumn(3).width = 14;
      orgSheet.getColumn(4).width = 14;
      orgSheet.getColumn(5).width = 14;
      orgSheet.getColumn(6).width = 14;

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const stamp = new Date().toISOString().slice(0, 10);
      a.href = url;
      a.download = `gubae-attendance-${stamp}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      if (err.response?.status === 401) logout();
      else
        setError(
          err.response?.data?.message ||
            err.message ||
            "Unable to download attendance Excel report",
        );
    }
  };

  const loadSummary = async () => {
    try {
      const r = await api.get("/bookings/attendance/summary");
      setSummary(r.data);
      setOrgPage(1);
    } catch (err) {
      if (err.response?.status === 401) logout();
      else
        setError(
          err.response?.data?.message || "Unable to load attendance summary",
        );
    }
  };

  const organizations = summary?.byOrganization || [];
  const ORG_PAGE_SIZE = 10;
  const totalOrgPages = Math.max(
    1,
    Math.ceil(organizations.length / ORG_PAGE_SIZE),
  );
  const paginatedOrganizations = organizations.slice(
    (orgPage - 1) * ORG_PAGE_SIZE,
    orgPage * ORG_PAGE_SIZE,
  );
  const orgStart = organizations.length ? (orgPage - 1) * ORG_PAGE_SIZE + 1 : 0;
  const orgEnd = Math.min(orgPage * ORG_PAGE_SIZE, organizations.length);

  const handleScan = async (decodedText) => {
    if (processingRef.current) return;
    processingRef.current = true;
    await stop();
    setError("");
    setResult(null);

    try {
      const r = await api.post("/bookings/attendance/scan", {
        qrData: decodedText,
      });
      playScanSound(true);
      setResult(r.data);
      await loadSummary();
    } catch (err) {
      playScanSound(false);
      vibrateRejected();
      setError(
        err.response?.data?.message ||
          "Unable to record attendance | ተሳትፎ መመዝገብ አልተቻለም",
      );
    } finally {
      setTimeout(() => {
        processingRef.current = false;
      }, 900);
    }
  };

  const start = async () => {
    setResult(null);
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) {
        audioContextRef.current = audioContextRef.current || new AudioCtx();
        if (audioContextRef.current.state === "suspended")
          await audioContextRef.current.resume();
      }
    } catch {}
    setError("");
    try {
      await stop();
      const scanner = new Html5Qrcode("qr-reader");
      scannerRef.current = scanner;
      await scanner.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: { width: 270, height: 270 }, aspectRatio: 1 },
        handleScan,
        () => {},
      );
      runningRef.current = true;
      setScanning(true);
    } catch (e) {
      setError(
        "Camera could not start. Allow camera permission and open the scanner using Phone (or computer). ካሜራው መጀመር አልቻለም! ስካነሩ የልክዎን ወይም የኮምፒውተርዎን ካሜራ እንዲጠቀም ፍቃድ ይስጡ",
      );
    }
  };

  useEffect(() => {
    loadSummary();
    return () => {
      stop();
      try {
        audioContextRef.current?.close();
      } catch {}
    };
  }, []);

  return (
    <div className="page">
      <header>
        <div>
          <h1>Gubae Attendance</h1>
          <span>Participant QR Scanner | የተሳታፊዎች QR ማንበቢያ </span>
        </div>
        <button className="ghost" onClick={logout}>
          <LogOut size={17} /> Logout
        </button>
      </header>

      <main>
        <section className="grid">
          <div className="card scanner">
            <div id="qr-reader"></div>
            {!scanning && (
              <button onClick={start}>
                <ScanLine size={20} /> Start Camera Scanner | ካሜራውን ያስጀምሩ
              </button>
            )}
            {scanning && (
              <button className="danger" onClick={stop}>
                <RefreshCw size={18} /> Stop Scanner | ምንባቡን ያስቁሙ
              </button>
            )}
            <p className="hint">
              Show the participant's QR code inside the square. The scanner
              reads the QR token and verifies it against the registered
              participants.|
            </p>
          </div>

          <div className="card result">
            {result?.participant ? (
              <>
                {result.alreadyCheckedIn ? (
                  <RefreshCw className="warningIcon" size={54} />
                ) : (
                  <CheckCircle2 className="successIcon" size={54} />
                )}
                <h2>
                  {result.alreadyCheckedIn ? "Already Present" : "Present ✓"}
                </h2>
                <div className="person">
                  <strong>{result.participant.name}</strong>
                  <span>{result.participant.organization}</span>
                  <span>{result.participant.sex}</span>
                </div>
                <p>{result.message}</p>
              </>
            ) : error ? (
              <>
                <XCircle
                  className="errorIcon "
                  // style={{ textDecoration: "red" }}
                  size={54}
                />
                <h2>Scan Not Accepted | ምንባብ ተቀባይነት አላገኘም </h2>
                <p>{error}</p>
              </>
            ) : (
              <>
                <Users size={48} />
                <h2>Ready to scan | ለምንባብ ዝግጁ ነው </h2>
                <p>
                  Scan a registered participant's QR code to mark them present.
                  | የተመዘገቡ ተሳታፊዎችን QR ኮድ አስነብብ እንደተገኙ ምልክት ለማድረግ እና ለመመዝገብ{" "}
                </p>
              </>
            )}
          </div>
        </section>

        <section className="card summary">
          <div className="summaryHead">
            <div>
              <h2>Live Attendance Summary | ጥቅላ የተገኙ አባላት ላይቭ</h2>
              <p className="summaryHint">
                Download the complete Present/Absent list after the event.
              </p>
            </div>
            <div className="summaryActions">
              <button className="small" onClick={loadSummary}>
                <RefreshCw size={15} /> Refresh
              </button>
              <button className="downloadBtn" onClick={downloadAttendanceList}>
                <Download size={15} /> Download Attendance List
              </button>
              <button
                className="resetBtn"
                onClick={() => {
                  setResetOpen(true);
                  setResetPhrase("");
                  setResetCount("");
                  setError("");
                }}
              >
                Reset Test Attendance{" "}
              </button>
            </div>
          </div>
          {summary && (
            <>
              <div className="stats">
                <div>
                  <b>{summary.totalRegistered}</b>
                  <span>Registered | የተመዘገቡ </span>
                </div>
                <div>
                  <b>{summary.totalPresent}</b>
                  <span>Present | የተገኙ </span>
                </div>
                <div>
                  <b>{summary.totalAbsent}</b>
                  <span>Absent | የቀሩ </span>
                </div>
                <div>
                  <b>{summary.men}</b>
                  <span>Men Present | የተገኙ ወንዶች </span>
                </div>
                <div>
                  <b>{summary.women}</b>
                  <span>Women Present | የተገኙ ሴቶች </span>
                </div>
              </div>
              <h3>Present by Organization | በድርጅት የተገኙት </h3>
              <div className="tableWrap">
                <table>
                  <thead>
                    <tr>
                      <th>Organization | ድርጅት </th>
                      <th>Present | የተገኙ </th>
                      <th>Men | ወንዶች </th>
                      <th>Women | ሴቶች </th>
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedOrganizations.map((r) => (
                      <tr key={r.organization}>
                        <td>{r.organization}</td>
                        <td>{r.total}</td>
                        <td>{r.men}</td>
                        <td>{r.women}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {organizations.length > 0 && (
                <div className="pagination">
                  <div className="paginationInfo">
                    Showing {orgStart}–{orgEnd} of {organizations.length}{" "}
                    organizations
                  </div>
                  <div className="paginationControls">
                    <button
                      className="small"
                      disabled={orgPage === 1}
                      onClick={() => setOrgPage((p) => Math.max(1, p - 1))}
                    >
                      Previous
                    </button>
                    {Array.from({ length: totalOrgPages }, (_, i) => i + 1).map(
                      (pageNumber) => (
                        <button
                          key={pageNumber}
                          className={`pageButton ${orgPage === pageNumber ? "active" : ""}`}
                          onClick={() => setOrgPage(pageNumber)}
                        >
                          {pageNumber}
                        </button>
                      ),
                    )}
                    <button
                      className="small"
                      disabled={orgPage === totalOrgPages}
                      onClick={() =>
                        setOrgPage((p) => Math.min(totalOrgPages, p + 1))
                      }
                    >
                      Next
                    </button>
                  </div>
                </div>
              )}
            </>
          )}
        </section>

        {resetOpen && (
          <div className="resetOverlay" role="dialog" aria-modal="true">
            <div className="resetModal">
              <h2>Reset Test Attendance</h2>
              <p>
                This permanently clears the current attendance check-ins for all
                participants. Use this only before the real event.
              </p>
              <div className="resetWarning">
                Current Present count:{" "}
                <strong>{summary?.totalPresent ?? 0}</strong>
              </div>
              <label>
                Type <strong>RESET ATTENDANCE</strong>
              </label>
              <input
                value={resetPhrase}
                onChange={(e) => setResetPhrase(e.target.value)}
                placeholder="RESET ATTENDANCE"
                autoComplete="off"
              />
              <label>
                Enter the current Present count:{" "}
                <strong>{summary?.totalPresent ?? 0}</strong>
              </label>
              <input
                value={resetCount}
                onChange={(e) =>
                  setResetCount(e.target.value.replace(/\D/g, ""))
                }
                inputMode="numeric"
                placeholder={String(summary?.totalPresent ?? 0)}
              />
              <div className="resetActions">
                <button
                  className="small"
                  disabled={resetBusy}
                  onClick={() => setResetOpen(false)}
                >
                  Cancel
                </button>
                <button
                  className="resetConfirm"
                  disabled={
                    resetBusy ||
                    resetPhrase !== "RESET ATTENDANCE" ||
                    String(resetCount) !== String(summary?.totalPresent ?? "")
                  }
                  onClick={resetAttendance}
                >
                  {resetBusy ? "Resetting…" : "Confirm Reset"}
                </button>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function logout() {
  localStorage.removeItem("adminToken");
  location.reload();
}

function App() {
  const [logged, setLogged] = useState(!!localStorage.getItem("adminToken"));
  return logged ? <Scanner /> : <Login onLogin={() => setLogged(true)} />;
}

createRoot(document.getElementById("root")).render(<App />);
