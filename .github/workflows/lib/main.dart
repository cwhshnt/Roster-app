import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:device_calendar/device_calendar.dart';
import 'package:timezone/data/latest.dart' as tz_data;
import 'package:timezone/timezone.dart' as tz;

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  tz_data.initializeTimeZones();
  runApp(const MaterialApp(
    debugShowCheckedModeBanner: false,
    home: RosterAppHome(),
  ));
}

class RosterAppHome extends StatefulWidget {
  const RosterAppHome({super.key});
  @override
  State<RosterAppHome> createState() => _RosterAppHomeState();
}

class _RosterAppHomeState extends State<RosterAppHome> {
  final DeviceCalendarPlugin _calendarPlugin = DeviceCalendarPlugin();
  String _statusLog = "準備就緒。請選取 Excel 檔案或使用下方快捷鍵。";
  String? _selectedCalendarId;
  DateTime _manualDate = DateTime.now();

  @override
  void initState() {
    super.initState();
    _initCalendar();
  }

  Future<void> _initCalendar() async {
    var permissionsGranted = await _calendarPlugin.hasPermissions();
    if (permissionsGranted.isSuccess && !permissionsGranted.data!) {
      permissionsGranted = await _calendarPlugin.requestPermissions();
    }
    final calendarsResult = await _calendarPlugin.findCalendars();
    if (calendarsResult.isSuccess && calendarsResult.data!.isNotEmpty) {
      setState(() {
        _selectedCalendarId = calendarsResult.data!.first.id;
      });
    }
  }

  // 1. 先 Del 後入鋼律
  Future<void> _clearOldEvents(DateTime start, DateTime end) async {
    if (_selectedCalendarId == null) return;
    final events = await _calendarPlugin.retrieveEvents(
      _selectedCalendarId,
      RetrieveEventsParams(startDate: start, endDate: end),
    );
    if (events.isSuccess && events.data != null) {
      for (var ev in events.data!) {
        final title = ev.title ?? '';
        if (title.contains('更') || title.contains('☀️') || title.contains('🗓️')) {
          await _calendarPlugin.deleteEvent(_selectedCalendarId, ev.eventId);
          await Future.delayed(const Duration(milliseconds: 50));
        }
      }
    }
  }

  // 2. 寫入單一更份（含 09:00-17:00 假期鋼律與 90 分鐘提醒）
  Future<void> _insertShift(DateTime date, String rawShift, {String desc = ''}) async {
    if (_selectedCalendarId == null) return;
    final clean = rawShift.replaceAll(RegExp(r'[#@◤+()]'), '').trim().toUpperCase();
    if (clean.isEmpty) return;

    final leaveTypes = ["AL", "WO", "O", "DO", "SH", "PHO", "SD", "SL", "OC1", "OC2", "CCLV"];
    Event event;

    if (leaveTypes.contains(clean)) {
      // 假期鎖定 09:00 - 17:00 單日不跨夜
      final start = DateTime(date.year, date.month, date.day, 9, 0);
      final end = DateTime(date.year, date.month, date.day, 17, 0);
      event = Event(
        _selectedCalendarId,
        title: '☀️ $clean',
        start: tz.TZDateTime.from(start, tz.local),
        end: tz.TZDateTime.from(end, tz.local),
        description: desc.isEmpty ? '👤 我的更份：【$clean】' : desc,
      );
    } else {
      int sH = 7, sM = 0, eH = 15, eM = 48;
      if (clean == 'A2') { eH = 14; eM = 36; }
      else if (clean == 'P') { sH = 12; sM = 45; eH = 21; eM = 33; }
      else if (clean == 'N') { sH = 21; sM = 10; eH = 23; eM = 59; }
      else if (clean == 'D') { sH = 8; sM = 45; eH = 17; eM = 33; }

      final start = DateTime(date.year, date.month, date.day, sH, sM);
      final end = DateTime(date.year, date.month, date.day, eH, eM);
      event = Event(
        _selectedCalendarId,
        title: '🗓️ $clean 更',
        start: tz.TZDateTime.from(start, tz.local),
        end: tz.TZDateTime.from(end, tz.local),
        description: desc.isEmpty ? '👤 我的更份：【$clean 更】' : desc,
        reminders: [Reminder(minutes: 90)],
      );
    }
    await _calendarPlugin.createOrUpdateEvent(event);
  }

  // 3. 讀取 Excel 並一鍵匯入
  Future<void> _pickAndProcessExcel() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xls', 'xlsx'],
    );
    if (result == null || result.files.single.path == null) return;

    setState(() => _statusLog = "正在解析更表...");
    final bytes = File(result.files.single.path!).readAsBytesSync();
    final excel = Excel.decodeBytes(bytes);
    final sheet = excel.tables[excel.tables.keys.first]!;

    // 尋找日期列與詹永康之行數
    int jamesRow = -1;
    for (int r = 0; r < sheet.maxRows; r++) {
      final rowData = sheet.row(r).map((c) => c?.value?.toString() ?? '').toList();
      if (rowData.any((s) => s.contains('詹永康') || s.contains('CHIM'))) {
        jamesRow = r;
        break;
      }
    }

    if (jamesRow == -1) {
      setState(() => _statusLog = "⚠️ 找不到「詹永康」更期資料，請檢查檔案！");
      return;
    }

    setState(() => _statusLog = "🧹 正在執行「先 Del 後入」清空舊日曆...");
    final now = DateTime.now();
    final startDate = DateTime(now.year, now.month, 1);
    final endDate = DateTime(now.year, now.month + 1, 5);
    await _clearOldEvents(startDate, endDate);

    setState(() => _statusLog = "⏳ 正在寫入新更期（極致防鎖緩衝中）...");
    // 依序寫入詹永康更份
    final rowShifts = sheet.row(jamesRow).skip(9).map((c) => c?.value?.toString() ?? '').where((s) => s.isNotEmpty && s != '\n').toList();
    for (int i = 0; i < rowShifts.length; i++) {
      final currentDay = startDate.add(Duration(days: i));
      await _insertShift(currentDay, rowShifts[i]);
      await Future.delayed(const Duration(milliseconds: 300));
    }

    setState(() => _statusLog = "✅ 匯入成功！整月更期已安全寫入 Google 日曆。");
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("1C 更表與日曆神助手"),
        backgroundColor: Colors.teal,
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            ElevatedButton.icon(
              style: ElevatedButton.styleFrom(
                padding: const EdgeInsets.symmetric(vertical: 16),
                backgroundColor: Colors.teal,
              ),
              onPressed: _pickAndProcessExcel,
              icon: const Icon(Icons.file_upload, color: Colors.white),
              label: const Text("📁 選取 Excel 檔一鍵匯入日曆", style: TextStyle(color: Colors.white, fontSize: 16)),
            ),
            const SizedBox(height: 16),
            Card(
              color: Colors.grey.shade100,
              child: Padding(
                padding: const EdgeInsets.all(12.0),
                child: Text(_statusLog, style: const TextStyle(fontSize: 14, height: 1.4)),
              ),
            ),
            const Divider(height: 32),
            const Text("📱 手動快捷改更面板", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
            const SizedBox(height: 8),
            Row(
              children: [
                Text("選擇日期: ${_manualDate.month}月${_manualDate.day}日"),
                const Spacer(),
                TextButton(
                  onPressed: () async {
                    final picked = await showDatePicker(
                      context: context,
                      initialDate: _manualDate,
                      firstDate: DateTime(2026, 1, 1),
                      lastDate: DateTime(2027, 12, 31),
                    );
                    if (picked != null) setState(() => _manualDate = picked);
                  },
                  child: const Text("更換日期"),
                ),
              ],
            ),
            const SizedBox(height: 8),
            // 人手 Button 快捷輸入面板
            Wrap(
              spacing: 8,
              runSpacing: 8,
              children: [
                _buildShiftBtn("A", Colors.green),
                _buildShiftBtn("A2", Colors.lightGreen),
                _buildShiftBtn("P", Colors.blue),
                _buildShiftBtn("N", Colors.grey),
                _buildShiftBtn("D", Colors.orange),
                _buildShiftBtn("WO", Colors.amber),
                _buildShiftBtn("O", Colors.amber.shade700),
                _buildShiftBtn("AL", Colors.deepOrange),
                _buildShiftBtn("SL", Colors.redAccent),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildShiftBtn(String shift, Color color) {
    return ElevatedButton(
      style: ElevatedButton.styleFrom(
        backgroundColor: color,
        foregroundColor: Colors.white,
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 12),
      ),
      onPressed: () async {
        setState(() => _statusLog = "正在更新 ${_manualDate.month}/${_manualDate.day} 為 $shift 更...");
        await _insertShift(_manualDate, shift);
        setState(() => _statusLog = "✅ ${_manualDate.month}/${_manualDate.day} 已成功改為 $shift 更！");
      },
      child: Text(shift, style: const TextStyle(fontWeight: FontWeight.bold)),
    );
  }
}
