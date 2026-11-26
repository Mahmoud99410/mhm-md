package com.mhm.md;

import android.content.ContentResolver;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.provider.OpenableColumns;
import android.util.Base64;
import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/** Receives files shared to the app ("Share" / "Open with") and hands them to the web layer. */
@CapacitorPlugin(name = "IncomingFile")
public class IncomingFilePlugin extends Plugin {
    private static final List<Uri> pending = new ArrayList<>();
    private static final List<String> pendingText = new ArrayList<>();
    private static IncomingFilePlugin instance;
    private static final long MAX_BYTES = 60L * 1024 * 1024;

    @Override
    public void load() {
        instance = this;
    }

    @SuppressWarnings("deprecation")
    static void handleIntent(Intent intent) {
        if (intent == null) return;
        String action = intent.getAction();
        boolean added = false;
        synchronized (pending) {
            if (Intent.ACTION_SEND.equals(action)) {
                Uri u = intent.getParcelableExtra(Intent.EXTRA_STREAM);
                if (u != null) { pending.add(u); added = true; }
                else {
                    String t = intent.getStringExtra(Intent.EXTRA_TEXT);
                    if (t != null && t.trim().length() > 0) { pendingText.add(t); added = true; }
                }
            } else if (Intent.ACTION_SEND_MULTIPLE.equals(action)) {
                ArrayList<Uri> list = intent.getParcelableArrayListExtra(Intent.EXTRA_STREAM);
                if (list != null) for (Uri u : list) if (u != null) { pending.add(u); added = true; }
            } else if (Intent.ACTION_VIEW.equals(action) && intent.getData() != null) {
                pending.add(intent.getData()); added = true;
            }
        }
        if (added && instance != null) {
            instance.notifyListeners("filesReceived", new JSObject(), true);
        }
    }

    @PluginMethod
    public void takePending(PluginCall call) {
        JSArray files = new JSArray();
        List<Uri> uris;
        List<String> texts;
        synchronized (pending) {
            uris = new ArrayList<>(pending);
            texts = new ArrayList<>(pendingText);
            pending.clear();
            pendingText.clear();
        }
        ContentResolver cr = getContext().getContentResolver();
        int n = 0;
        for (Uri uri : uris) {
            try {
                String name = displayName(cr, uri);
                InputStream in = cr.openInputStream(uri);
                if (in == null) continue;
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                byte[] buf = new byte[65536];
                int r;
                long total = 0;
                while ((r = in.read(buf)) != -1) {
                    total += r;
                    if (total > MAX_BYTES) break;
                    out.write(buf, 0, r);
                }
                in.close();
                if (total > MAX_BYTES) continue;
                JSObject f = new JSObject();
                f.put("name", name);
                f.put("data", Base64.encodeToString(out.toByteArray(), Base64.NO_WRAP));
                files.put(f);
            } catch (Exception e) {
                // skip unreadable file
            }
        }
        for (String t : texts) {
            JSObject f = new JSObject();
            f.put("name", "shared-text-" + (++n) + ".txt");
            f.put("data", Base64.encodeToString(t.getBytes(java.nio.charset.StandardCharsets.UTF_8), Base64.NO_WRAP));
            files.put(f);
        }
        JSObject ret = new JSObject();
        ret.put("files", files);
        call.resolve(ret);
    }

    private static String displayName(ContentResolver cr, Uri uri) {
        String name = null;
        try (Cursor c = cr.query(uri, new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
            if (c != null && c.moveToFirst()) name = c.getString(0);
        } catch (Exception e) { /* ignore */ }
        if (name == null) name = uri.getLastPathSegment();
        if (name == null) name = "file";
        if (!name.contains(".")) {
            String type = cr.getType(uri);
            if (type != null) {
                if (type.contains("pdf")) name += ".pdf";
                else if (type.contains("wordprocessingml")) name += ".docx";
                else if (type.contains("presentationml")) name += ".pptx";
                else if (type.contains("spreadsheetml")) name += ".xlsx";
                else if (type.startsWith("image/")) name += "." + type.substring(6).replace("jpeg", "jpg");
                else if (type.startsWith("text/")) name += ".txt";
            }
        }
        return name;
    }
}
