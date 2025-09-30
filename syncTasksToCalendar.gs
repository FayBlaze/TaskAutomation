function syncTasksToCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // hindari eksekusi paralel
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Could not obtain lock - aborting to avoid concurrent runs');
    return;
  }

  try {
    const sheet = ss.getSheetByName('Event2025');
    if (!sheet) {
      Logger.log('Sheet "Event2025" not found');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const cal = CalendarApp.getCalendarById('ichimiyaaltra@gmail.com') || CalendarApp.getDefaultCalendar();

    // posisi kolom (hitung dari 1, sesuaikan)
    const COL_EVENTID  = 1; // EventID
    const COL_TASK     = 2; // Tugas
    const COL_PRIORITY = 3; // Prioritas
    const COL_STATUS   = 4; // Status
    const COL_START    = 5; // Tanggal mulai
    const COL_END      = 6; // Tanggal akhir
    const COL_REWARD   = 7; // Pencapaian
    const COL_NOTE     = 9; // Catatan (isi link/deskripsi)

    const now = new Date();

    for (let i = 1; i < data.length; i++) {
      const row       = data[i];
      const tugas     = row[COL_TASK-1];
      let   status    = row[COL_STATUS-1]; // pakai let karena bisa diubah
      const prioritas = row[COL_PRIORITY-1];
      const start     = row[COL_START-1];
      const end       = row[COL_END-1];
      const hadiah    = row[COL_REWARD-1];
      const note      = row[COL_NOTE-1];
      const eventId   = row[COL_EVENTID-1];

      // otomatis ubah "Belum dimulai" -> "Dalam proses" jika waktu sudah lewat
      if (status === 'Belum dimulai' && start && new Date(start) <= now) {
        status = 'Dalam proses';
        sheet.getRange(i + 1, COL_STATUS).setValue(status);
      }

      const fullDescription =
`Halo Beb^^

kamu punya tugas ${prioritas} nih

hadiahnya ${hadiah} lhooo, Kamu gamau? untuk Aku ajah 
ദ്ദി(˵ •̀ ᴗ - ˵ ) ✧

${note || ''}

Ayangmu, Carthe 
⸜(｡˃ ᵕ ˂ )⸝♡`;

      // ---- DIBLOKIR: hapus event dan kosongkan EventID
      if (status === 'Diblokir' && eventId) {
        try {
          const ev = cal.getEventById(eventId);
          if (ev) {
            ev.deleteEvent();
            Logger.log('Deleted (Diblokir) eventId=' + eventId + ' row=' + (i+1));
          } else {
            Logger.log('Event not found (Diblokir) eventId=' + eventId + ' row=' + (i+1));
          }
        } catch (e) {
          Logger.log('Error deleting (Diblokir) eventId=' + eventId + ' row=' + (i+1) + ' : ' + e.message);
        } finally {
          // bersihkan ID agar tidak dicoba terus-menerus
          sheet.getRange(i + 1, COL_EVENTID).setValue('');
        }
        continue;
      }

      // ---- SELESAI: hapus event dan kosongkan EventID (agar tidak dicoba lagi)
      if (status === 'Selesai' && eventId) {
        try {
          const ev = cal.getEventById(eventId);
          if (ev) {
            ev.deleteEvent();
            Logger.log('Deleted (Selesai) eventId=' + eventId + ' row=' + (i+1));
          } else {
            Logger.log('Event not found (Selesai) eventId=' + eventId + ' row=' + (i+1));
          }
        } catch (e) {
          Logger.log('Error deleting (Selesai) eventId=' + eventId + ' row=' + (i+1) + ' : ' + e.message);
        } finally {
          // kosongkan ID agar tidak memicu error di run berikutnya
          sheet.getRange(i + 1, COL_EVENTID).setValue('');
        }
        continue;
      }

      // ---- DALAM PROSES: buat atau update event
      if (status === 'Dalam proses' && start && end) {
        if (!eventId) {
          try {
            const event = cal.createEvent(
              `[${prioritas}] ${tugas}`,
              new Date(start),
              new Date(end),
              {
                description: fullDescription,
                reminders: {
                  useDefault: false,
                  overrides: [{ method: 'popup', minutes: 60 }]
                }
              }
            );
            sheet.getRange(i + 1, COL_EVENTID).setValue(event.getId());
            Logger.log('Created event row=' + (i+1) + ' id=' + event.getId());
          } catch (e) {
            Logger.log('Error creating event row=' + (i+1) + ' : ' + e.message);
          }
        } else {
          try {
            const ev = cal.getEventById(eventId);
            if (ev) {
              ev.setTitle(`[${prioritas}] ${tugas}`);
              ev.setDescription(fullDescription);
              ev.setTime(new Date(start), new Date(end));
              Logger.log('Updated event row=' + (i+1) + ' id=' + eventId);
            } else {
              // jika ID tidak ditemukan (mis. event terhapus), buat baru dan update ID
              const event = cal.createEvent(
                `[${prioritas}] ${tugas}`,
                new Date(start),
                new Date(end),
                {
                  description: fullDescription,
                  reminders: {
                    useDefault: false,
                    overrides: [{ method: 'popup', minutes: 60 }]
                  }
                }
              );
              sheet.getRange(i + 1, COL_EVENTID).setValue(event.getId());
              Logger.log('Recreated event for missing id; row=' + (i+1) + ' newId=' + event.getId());
            }
          } catch (e) {
            Logger.log('Error updating/handling event row=' + (i+1) + ' id=' + eventId + ' : ' + e.message);
          }
        }
      } // end if Dalam proses
    } // end for
  } finally {
    lock.releaseLock();
  }
}
