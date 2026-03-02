unit DataGrid;

interface

uses System, System.Drawing, System.Windows.Forms, System.Reflection, System.IO, System.Text;

type
  TGridType = (gtNormal, gtItem, gtMap, gtMob, gtSkill, gtNpc, gtMorph,gtFamiliar,gtDamageSkin,gtReactor);
  
  DataViewer = class(DataGridView)
    constructor Create(GridType: TGridType);
    begin
      Self.Dock := System.Windows.Forms.DockStyle.Fill;
      RowTemplate.Height := 80;
      
      //Self.AutoSizeColumnsMode := DataGridViewAutoSizeColumnsMode.Fill;
      var ID := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var Icon := new System.Windows.Forms.DataGridViewImageColumn;
      var MorphIcon := new System.Windows.Forms.DataGridViewImageColumn;
     
      var Name := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var Info := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var Desc := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var Level := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var MapName := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var StreetName := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var Speak := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var MorphID := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var FamiliarSkillName := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var FamiliarSkillDesc := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var FamiliarCardID := new System.Windows.Forms.DataGridViewTextBoxColumn;
      var FamiliarIcon := new System.Windows.Forms.DataGridViewImageColumn;
      var Sample := new System.Windows.Forms.DataGridViewImageColumn;
      //  Self.ColumnHeadersHeightSizeMode := System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      Self.DefaultCellStyle.Font := new System.Drawing.Font('Segoe UI', 8.5);  
      ID.DataPropertyName := 'ID';
      iD.HeaderText := 'ID';
      iD.Name := 'propID';
      ID.ReadOnly := true;
      id.DefaultCellStyle.Alignment := DataGridViewContentAlignment.MiddleCenter;
      ID.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter;      
      ID.Width := 90;   
      
      MorphID.DataPropertyName := 'MorphID';
      MorphID.HeaderText := 'MorphID';
      MorphID.Name := 'propID';
      MorphID.ReadOnly := true;
      MorphID.DefaultCellStyle.Alignment := DataGridViewContentAlignment.MiddleCenter;
      MorphID.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter;      
      MorphID.Width := 90;   
      
      Icon.DataPropertyName := 'Icon';
      Icon.HeaderText := 'Icon';
      Icon.Name := 'propBitmap';
      Icon.ReadOnly := true;
      Icon.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Icon.Width := 100;
      
      MorphIcon.DataPropertyName := 'MorphIcon';
      MorphIcon.HeaderText := 'MorphIcon';
      MorphIcon.Name := 'propBitmap';
      MorphIcon.ReadOnly := true;
      MorphIcon.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      MorphIcon.Width := 100;
      
      FamiliarIcon.DataPropertyName := 'FamiliarIcon';
      FamiliarIcon.HeaderText := 'Familiar';
      FamiliarIcon.Name := 'propBitmap';
      FamiliarIcon.ReadOnly := true;
      FamiliarIcon.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      FamiliarIcon.Width := 100;
      
      Sample.DataPropertyName := 'Sample';
      Sample.HeaderText := 'Sample';
      Sample.Name := 'propBitmap';
      Sample.ReadOnly := true;
      Sample.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Sample.Width := 280;
      
      Name.DataPropertyName := 'NameProperty';
      Name.HeaderText := 'Name';
      Name.Name := 'propName';
      Name.ReadOnly := true;
      Name.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Name.Width := 200;
      
      FamiliarSkillName.DataPropertyName := 'NameProperty';
      FamiliarSkillName.HeaderText := 'Skill Name';
      FamiliarSkillName.Name := 'propName';
      FamiliarSkillName.ReadOnly := true;
      FamiliarSkillName.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      FamiliarSkillName.Width := 100;
      
      FamiliarSkillDesc.DataPropertyName := 'NameProperty';
      FamiliarSkillDesc.HeaderText := 'Skill Desc';
      FamiliarSkillDesc.Name := 'propName';
      FamiliarSkillDesc.ReadOnly := true;
      FamiliarSkillDesc.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      FamiliarSkillDesc.Width := 100;
      
      FamiliarCardID.DataPropertyName := 'NameProperty';
      FamiliarCardID.HeaderText := 'Card ID';
      FamiliarCardID.Name := 'propName';
      FamiliarCardID.ReadOnly := true;
      FamiliarCardID.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      FamiliarCardID.Width := 100;
      
      Desc.DataPropertyName := 'DescProperty';
      Desc.HeaderText := 'Description';
      Desc.Name := 'Desc';
      Desc.ReadOnly := true;
      Desc.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Desc.Width := 100;
      
      Level.DataPropertyName := 'LevelProperty';
      Level.HeaderText := 'Level';
      Level.Name := 'Level';
      Level.ReadOnly := true;
      Level.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Level.Width := 100;
      
      StreetName.DataPropertyName := 'StreetNameProperty';
      StreetName.HeaderText := 'StreetName';
      StreetName.Name := 'StreetName';
      StreetName.ReadOnly := True;
      StreetName.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      StreetName.Width := 100;
      
      MapName.DataPropertyName := 'MapNameProperty';
      MapName.HeaderText := 'MapName';
      MapName.Name := 'MapName';
      Mapname.ReadOnly := True;
      MapName.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      MapName.Width := 100;
      
      Speak.DataPropertyName := 'SpeakNameProperty';
      Speak.HeaderText := 'Speak';
      Speak.Name := 'Speak';
      Speak.ReadOnly := True;
      Speak.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Speak.Width := 800;  
      
      Info.DataPropertyName := 'PropertiesProperty';
      Info.HeaderText := 'Info';
      Info.Name := 'propProperties';
      Info.ReadOnly := true;
      Info.Width := 350;
      Info.HeaderCell.Style.Alignment := DataGridViewContentAlignment.MiddleCenter; 
      Info.AutoSizeMode := System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
      DefaultCellStyle.WrapMode := DataGridViewTriState.True;
       ColumnHeadersHeight := 28; 
      case GridType of
        gtNormal:
          begin
            RowTemplate.Height := 40;
            Desc.Width := 600;
            Info.Width := 200;
           
            Self.Columns.AddRange(ID, Icon, Name, Info);
          end;
        gtItem: 
          begin
            RowTemplate.Height := 40;
            Desc.Width := 600;
            Info.Width := 200;
           
            Self.Columns.AddRange(ID, icon, Name, Desc, Info);
          end;
        gtMap: 
          begin
            RowTemplate.Height := 60;
            Desc.Width := 600;
            Info.Width := 200;
            StreetName.Width:=200;
            MapName.Width:=200;
            Self.Columns.AddRange(ID, Icon, StreetName, MapName, Info);
          end;
        gtMob: 
          begin
            RowTemplate.Height := 80;
            Icon.Width := 150;
            Desc.Width := 600;
            Info.Width := 200;
           
            Self.Columns.AddRange(ID, Icon, Name, Info);
          end;
        gtSkill: 
          begin
            RowTemplate.Height := 60;
            Icon.Width := 80;
            Desc.Width := 400;
            Info.Width := 200;
            Level.width:=450;
            Self.Columns.AddRange(ID, Icon, Name, Desc, Level, Info);
          end; 
        gtNpc:
          begin
            RowTemplate.Height := 70;
            Desc.Width := 600;
            Info.Width := 200;
            Self.Columns.AddRange(ID, Icon, Name, Speak, Info);
          end;
        gtMorph:
          begin
            RowTemplate.Height := 70;
            Desc.Width := 600;
            Self.Columns.AddRange(ID, Icon, MorphID, MorphIcon, Name, Desc, Info);
          end;       
        gtFamiliar:
          begin
            RowTemplate.Height := 70;
            ColumnHeadersHeight := 28; 
            Desc.Width := 300;
            FamiliarSkillName.Width:=150;
            FamiliarSkillDesc.Width:=350;
            Info.AutoSizeMode := System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            Name.AutoSizeMode := System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            Info.Width := 300;
            Self.Columns.AddRange(ID, FamiliarIcon,Info,FamiliarSkillName,FamiliarSkillDesc,FamiliarCardID,Icon, Name);
          end;
          gtDamageSkin:
          begin
            RowTemplate.Height := 50;
            ColumnHeadersHeight := 28; 
            Name.Width:=250;
            Desc.Width := 600;
            Desc.AutoSizeMode := System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            Self.Columns.AddRange(ID, Icon, Name,Sample,Desc);
          end;  
          gtReactor:
          begin
            RowTemplate.Height := 80;
            Info.Width := 300;
            Self.Columns.AddRange(ID, Icon, Info);
          end;
         
      end;
      if (not System.Windows.Forms.SystemInformation.TerminalServerSession) then
      begin
        var dgvType := Self.GetType();
        var pi: PropertyInfo := dgvType.GetProperty('DoubleBuffered', BindingFlags.Instance or BindingFlags.NonPublic);
        pi.SetValue(Self, True, nil);
      end;
    end;
  end;

implementation

function JsonEscape(const S: string): string;
begin
  if S = '' then
  begin
    Result := '';
    Exit;
  end;
  
  var B := new StringBuilder;
  foreach var Ch in S do
  begin
    case Ch of
      '"': B.Append('\"');
      '\': B.Append('\\');
      #8: B.Append('\b');
      #9: B.Append('\t');
      #10: B.Append('\n');
      #12: B.Append('\f');
      #13: B.Append('\r');
    else
      if Ord(Ch) < 32 then
        B.Append('\u' + Ord(Ch).ToString('X4'))
      else
        B.Append(Ch);
    end;
  end;
  Result := B.ToString;
end;

function SanitizeName(const S: string): string;
begin
  var B := new StringBuilder;
  foreach var Ch in S do
  begin
    if Char.IsLetterOrDigit(Ch) then
      B.Append(Ch)
    else if (Ch = '-') or (Ch = '_') then
      B.Append(Ch)
    else
      B.Append('_');
  end;
  
  Result := B.ToString;
  while (Result <> '') and (Result[1] = '_') do
    Delete(Result, 1, 1);
  while (Result <> '') and (Result[Length(Result)] = '_') do
    Delete(Result, Length(Result), 1);
  if Result = '' then
    Result := 'tab';
end;

function BuildWebIndexHtml: string;
begin
  var Html := new StringBuilder;
  Html.AppendLine('<!doctype html>');
  Html.AppendLine('<html lang="en">');
  Html.AppendLine('<head>');
  Html.AppendLine('  <meta charset="utf-8">');
  Html.AppendLine('  <meta name="viewport" content="width=device-width,initial-scale=1">');
  Html.AppendLine('  <title>GMHandTool WZ Export</title>');
  Html.AppendLine('  <style>');
  Html.AppendLine('    :root { --bg:#f4f5f7; --panel:#ffffff; --ink:#1f2937; --muted:#6b7280; --line:#d1d5db; --brand:#0f766e; }');
  Html.AppendLine('    * { box-sizing: border-box; }');
  Html.AppendLine('    body { margin:0; font-family: "Segoe UI", Tahoma, sans-serif; color:var(--ink); background:linear-gradient(180deg,#f7f7f8,#eef2f7); }');
  Html.AppendLine('    .top { position:sticky; top:0; z-index:2; background:var(--panel); border-bottom:1px solid var(--line); padding:12px; display:flex; gap:8px; flex-wrap:wrap; align-items:center; }');
  Html.AppendLine('    input, select { height:36px; border:1px solid var(--line); border-radius:8px; padding:0 10px; font-size:14px; }');
  Html.AppendLine('    input { min-width:260px; flex:1; }');
  Html.AppendLine('    .meta { color:var(--muted); font-size:13px; margin-left:auto; }');
  Html.AppendLine('    #results { padding:12px; display:grid; grid-template-columns:repeat(auto-fill,minmax(340px,1fr)); gap:10px; }');
  Html.AppendLine('    .card { display:grid; grid-template-columns:84px 1fr; gap:10px; background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:10px; }');
  Html.AppendLine('    .thumb { width:72px; height:72px; border:1px solid var(--line); border-radius:8px; object-fit:contain; background:#fff; }');
  Html.AppendLine('    h3 { margin:0 0 6px 0; font-size:14px; color:var(--brand); }');
  Html.AppendLine('    p { margin:2px 0; font-size:12px; line-height:1.35; }');
  Html.AppendLine('    .empty { padding:20px; color:var(--muted); }');
  Html.AppendLine('  </style>');
  Html.AppendLine('  <script src="manifest.js"></script>');
  Html.AppendLine('</head>');
  Html.AppendLine('<body>');
  Html.AppendLine('  <div class="top">');
  Html.AppendLine('    <input id="q" type="text" placeholder="Search ID, Name, Info...">');
  Html.AppendLine('    <select id="tabFilter"><option value="">All tabs</option></select>');
  Html.AppendLine('    <div id="count" class="meta">0 records</div>');
  Html.AppendLine('  </div>');
  Html.AppendLine('  <div id="results"></div>');
  Html.AppendLine('  <script>');
  Html.AppendLine('    const state = { rows: [], query: "", tab: "" };');
  Html.AppendLine('    function loadScript(src) {');
  Html.AppendLine('      return new Promise((resolve, reject) => {');
  Html.AppendLine('        const s = document.createElement("script");');
  Html.AppendLine('        s.src = src;');
  Html.AppendLine('        s.onload = () => resolve();');
  Html.AppendLine('        s.onerror = () => reject(new Error("failed: " + src));');
  Html.AppendLine('        document.head.appendChild(s);');
  Html.AppendLine('      });');
  Html.AppendLine('    }');
  Html.AppendLine('    function firstImagePath(images) {');
  Html.AppendLine('      if (!images) return "";');
  Html.AppendLine('      const keys = Object.keys(images);');
  Html.AppendLine('      return keys.length > 0 ? images[keys[0]] : "";');
  Html.AppendLine('    }');
  Html.AppendLine('    function applyFilter() {');
  Html.AppendLine('      const q = state.query.toLowerCase();');
  Html.AppendLine('      return state.rows.filter((r) => {');
  Html.AppendLine('        if (state.tab && r.tab !== state.tab) return false;');
  Html.AppendLine('        if (!q) return true;');
  Html.AppendLine('        const hay = ((r.search || "") + " " + (r.id || "") + " " + (r.name || "")).toLowerCase();');
  Html.AppendLine('        return hay.includes(q);');
  Html.AppendLine('      });');
  Html.AppendLine('    }');
  Html.AppendLine('    function render() {');
  Html.AppendLine('      const root = document.getElementById("results");');
  Html.AppendLine('      root.innerHTML = "";');
  Html.AppendLine('      const filtered = applyFilter();');
  Html.AppendLine('      document.getElementById("count").textContent = filtered.length + " records";');
  Html.AppendLine('      if (filtered.length === 0) {');
  Html.AppendLine('        const empty = document.createElement("div");');
  Html.AppendLine('        empty.className = "empty";');
  Html.AppendLine('        empty.textContent = "No results for this filter.";');
  Html.AppendLine('        root.appendChild(empty);');
  Html.AppendLine('        return;');
  Html.AppendLine('      }');
  Html.AppendLine('      const maxRows = 500;');
  Html.AppendLine('      filtered.slice(0, maxRows).forEach((row) => {');
  Html.AppendLine('        const card = document.createElement("article");');
  Html.AppendLine('        card.className = "card";');
  Html.AppendLine('        const img = document.createElement("img");');
  Html.AppendLine('        img.className = "thumb";');
  Html.AppendLine('        img.src = row.preview || firstImagePath(row.images);');
  Html.AppendLine('        if (!img.src) img.style.visibility = "hidden";');
  Html.AppendLine('        const body = document.createElement("div");');
  Html.AppendLine('        const title = document.createElement("h3");');
  Html.AppendLine('        title.textContent = row.tab + " | " + (row.id || "-") + " | " + (row.name || "-");');
  Html.AppendLine('        body.appendChild(title);');
  Html.AppendLine('        const fields = row.fields || {};');
  Html.AppendLine('        const keys = Object.keys(fields).slice(0, 8);');
  Html.AppendLine('        keys.forEach((k) => {');
  Html.AppendLine('          const p = document.createElement("p");');
  Html.AppendLine('          p.textContent = k + ": " + fields[k];');
  Html.AppendLine('          body.appendChild(p);');
  Html.AppendLine('        });');
  Html.AppendLine('        card.appendChild(img);');
  Html.AppendLine('        card.appendChild(body);');
  Html.AppendLine('        root.appendChild(card);');
  Html.AppendLine('      });');
  Html.AppendLine('      if (filtered.length > maxRows) {');
  Html.AppendLine('        const info = document.createElement("div");');
  Html.AppendLine('        info.className = "empty";');
  Html.AppendLine('        info.textContent = "Showing first " + maxRows + " results.";');
  Html.AppendLine('        root.appendChild(info);');
  Html.AppendLine('      }');
  Html.AppendLine('    }');
  Html.AppendLine('    function buildTabFilter() {');
  Html.AppendLine('      const sel = document.getElementById("tabFilter");');
  Html.AppendLine('      const tabs = Array.from(new Set(state.rows.map((r) => r.tab))).sort();');
  Html.AppendLine('      tabs.forEach((tab) => {');
  Html.AppendLine('        const opt = document.createElement("option");');
  Html.AppendLine('        opt.value = tab;');
  Html.AppendLine('        opt.textContent = tab;');
  Html.AppendLine('        sel.appendChild(opt);');
  Html.AppendLine('      });');
  Html.AppendLine('    }');
  Html.AppendLine('    function bindEvents() {');
  Html.AppendLine('      document.getElementById("q").addEventListener("input", (e) => {');
  Html.AppendLine('        state.query = e.target.value || "";');
  Html.AppendLine('        render();');
  Html.AppendLine('      });');
  Html.AppendLine('      document.getElementById("tabFilter").addEventListener("change", (e) => {');
  Html.AppendLine('        state.tab = e.target.value || "";');
  Html.AppendLine('        render();');
  Html.AppendLine('      });');
  Html.AppendLine('    }');
  Html.AppendLine('    async function boot() {');
  Html.AppendLine('      const manifest = window.GM_MANIFEST || [];');
  Html.AppendLine('      for (const src of manifest) await loadScript(src);');
  Html.AppendLine('      const all = window.GM_EXPORTS || {};');
  Html.AppendLine('      Object.keys(all).forEach((k) => {');
  Html.AppendLine('        const rows = all[k] || [];');
  Html.AppendLine('        rows.forEach((r) => state.rows.push(r));');
  Html.AppendLine('      });');
  Html.AppendLine('      buildTabFilter();');
  Html.AppendLine('      bindEvents();');
  Html.AppendLine('      render();');
  Html.AppendLine('    }');
  Html.AppendLine('    boot().catch((err) => {');
  Html.AppendLine('      const root = document.getElementById("results");');
  Html.AppendLine('      root.innerHTML = "<div class=""empty"">Failed to load data: " + err.message + "</div>";');
  Html.AppendLine('    });');
  Html.AppendLine('  </script>');
  Html.AppendLine('</body>');
  Html.AppendLine('</html>');
  Result := Html.ToString;
end;

procedure RebuildManifest(BaseDir: string);
begin
  var Scripts := new List<string>;
  foreach var FileName in Directory.GetFiles(BaseDir, 'data.js', SearchOption.AllDirectories) do
  begin
    var RelPath := FileName.Substring(BaseDir.Length).TrimStart('\');
    Scripts.Add(RelPath.Replace('\', '/'));
  end;
  
  Scripts.Sort;
  var Manifest := new StringBuilder;
  Manifest.Append('window.GM_MANIFEST = [');
  for var i := 0 to Scripts.Count - 1 do
  begin
    if i > 0 then
      Manifest.Append(',');
    Manifest.Append('"').Append(JsonEscape(Scripts[i])).Append('"');
  end;
  Manifest.Append('];');
  
  System.IO.File.WriteAllText(System.IO.Path.Combine(BaseDir, 'manifest.js'), Manifest.ToString, Encoding.UTF8);
  System.IO.File.WriteAllText(System.IO.Path.Combine(BaseDir, 'index.html'), BuildWebIndexHtml(), Encoding.UTF8);
end;

procedure SaveBin(Self: DataViewer; Path: string); extensionmethod;
begin
  var BinWriter: BinaryWriter := new BinaryWriter(System.IO.File.Open(Path, FileMode.Create));
  BinWriter.Write(Self.Columns.Count);
  BinWriter.Write(Self.Rows.Count);
  foreach var Row: DataGridViewRow in Self.Rows do
  begin
    for var j := 0 to Self.Columns.Count - 1 do
    begin
      var val: object := Row.Cells[j].Value;
      if val is string then
      begin
        BinWriter.Write('str');
        BinWriter.Write(val.ToString);
      end 
      else 
      if val is Bitmap then
      begin
        BinWriter.Write('Bitmap');
        var MemStream: MemoryStream := new MemoryStream;
        Bitmap(Val).Save(MemStream, System.Drawing.Imaging.ImageFormat.Png);
        var Buffer: array of Byte := MemStream.GetBuffer;
        BinWriter.Write(Buffer.Length);
        BinWriter.Write(Buffer);
      end;
    end;
  end;
  BinWriter.Close;
end;

procedure LoadBin(Self: DataViewer; Path: string); extensionmethod;
begin
  var BinReader: BinaryReader := new BinaryReader(System.IO.File.Open(Path, FileMode.Open));
  var n: Integer := BinReader.ReadInt32;
  var m: Integer := BinReader.ReadInt32;
  for var i := 0 to m - 2 do
  begin
    Self.Rows.Add;
    for var j := 0 to n - 1 do
    begin
      if BinReader.ReadString = 'str' then
        Self.Rows[i].Cells[j].Value := BinReader.ReadString
      else 
      begin
        var BufferLength: Integer := BinReader.ReadInt32;
        var Buffer: array of Byte := BinReader.ReadBytes(BufferLength);
        var MemStream: MemoryStream := new MemoryStream(Buffer);
        var LImage: Image := Image.FromStream(MemStream);
        var Bmp: Bitmap := new Bitmap(LImage.Width, LImage.Height, System.Drawing.Imaging.PixelFormat.Format16bppRgb555);
        Bmp.MakeTransparent;
        var g := Graphics.FromImage(Bmp);
        g.DrawImage(LImage, new Rectangle(0, 0, LImage.Width, LImage.Height));
        Self.Rows[i].Cells[j].Value := Bmp; 
      end;
    end;
  end;
  BinReader.Close;
end;

function ExportWeb(Self: DataViewer; BaseDir, TabName: string): Integer; extensionmethod;
begin
  if not System.IO.Directory.Exists(BaseDir) then
    System.IO.Directory.CreateDirectory(BaseDir);
  
  var TabFolder := SanitizeName(TabName);
  var TabDir := System.IO.Path.Combine(BaseDir, TabFolder);
  var ImageDir := System.IO.Path.Combine(TabDir, 'images');
  
  if not System.IO.Directory.Exists(TabDir) then
    System.IO.Directory.CreateDirectory(TabDir);
  
  if System.IO.Directory.Exists(ImageDir) then
  begin
    foreach var OldPng in System.IO.Directory.GetFiles(ImageDir, '*.png') do
      System.IO.File.Delete(OldPng);
  end
  else
    System.IO.Directory.CreateDirectory(ImageDir);
  
  var RowsJson := new StringBuilder;
  RowsJson.Append('[');
  
  var HasRows := False;
  var ExportedRows := 0;
  var RowIndex := 0;
  
  foreach var Row: DataGridViewRow in Self.Rows do
  begin
    if Row.IsNewRow then
      continue;
    
    var FieldsJson := new StringBuilder;
    FieldsJson.Append('{');
    var HasField := False;
    
    var ImagesJson := new StringBuilder;
    ImagesJson.Append('{');
    var HasImage := False;
    
    var SearchText := new StringBuilder;
    var IDValue := '';
    var NameValue := '';
    var FirstText := '';
    var PreviewPath := '';
    
    for var j := 0 to Self.Columns.Count - 1 do
    begin
      var ColName := Self.Columns[j].HeaderText;
      var ColKey := ColName.ToLower;
      var CellValue: object := Row.Cells[j].Value;
      if CellValue = nil then
        continue;
      
      if (CellValue is Bitmap) or (CellValue is Image) then
      begin
        var ImgFileName := RowIndex.ToString('D6') + '_' + j.ToString('D2') + '_' + SanitizeName(ColName) + '.png';
        var ImgFullPath := System.IO.Path.Combine(ImageDir, ImgFileName);
        var ImgRelativePath := TabFolder + '/images/' + ImgFileName;
        
        if CellValue is Bitmap then
          Bitmap(CellValue).Save(ImgFullPath, System.Drawing.Imaging.ImageFormat.Png)
        else
        begin
          var TmpBmp := new Bitmap(Image(CellValue));
          TmpBmp.Save(ImgFullPath, System.Drawing.Imaging.ImageFormat.Png);
          TmpBmp.Dispose;
        end;
        
        if HasImage then
          ImagesJson.Append(',')
        else
          HasImage := True;
        ImagesJson.Append('"').Append(JsonEscape(ColName)).Append('":"').Append(JsonEscape(ImgRelativePath)).Append('"');
        
        if PreviewPath = '' then
          PreviewPath := ImgRelativePath;
      end
      else
      begin
        var TextValue := CellValue.ToString;
        
        if HasField then
          FieldsJson.Append(',')
        else
          HasField := True;
        FieldsJson.Append('"').Append(JsonEscape(ColName)).Append('":"').Append(JsonEscape(TextValue)).Append('"');
        
        if FirstText = '' then
          FirstText := TextValue;
        
        if (TextValue <> '') then
          SearchText.Append(ColName).Append(':').Append(TextValue).Append(' ');
        
        if (IDValue = '') and ((ColKey = 'id') or (ColKey = 'morphid') or (ColKey = 'card id')) then
          IDValue := TextValue;
        
        if (NameValue = '') and ((ColKey = 'name') or (ColKey = 'mapname') or (ColKey = 'streetname') or (ColKey =
          'skill name')) then
          NameValue := TextValue;
      end;
    end;
    
    FieldsJson.Append('}');
    ImagesJson.Append('}');
    
    if IDValue = '' then
      IDValue := FirstText;
    if NameValue = '' then
      NameValue := FirstText;
    
    if HasRows then
      RowsJson.Append(',')
    else
      HasRows := True;
    
    RowsJson.Append('{');
    RowsJson.Append('"tab":"').Append(JsonEscape(TabName)).Append('",');
    RowsJson.Append('"id":"').Append(JsonEscape(IDValue)).Append('",');
    RowsJson.Append('"name":"').Append(JsonEscape(NameValue)).Append('",');
    RowsJson.Append('"search":"').Append(JsonEscape(SearchText.ToString.Trim)).Append('",');
    RowsJson.Append('"fields":').Append(FieldsJson.ToString).Append(',');
    RowsJson.Append('"images":').Append(ImagesJson.ToString).Append(',');
    RowsJson.Append('"preview":"').Append(JsonEscape(PreviewPath)).Append('"');
    RowsJson.Append('}');
    
    ExportedRows += 1;
    RowIndex += 1;
  end;
  
  RowsJson.Append(']');
  
  System.IO.File.WriteAllText(System.IO.Path.Combine(TabDir, 'data.json'), RowsJson.ToString, Encoding.UTF8);
  System.IO.File.WriteAllText(System.IO.Path.Combine(TabDir, 'data.js'), 'window.GM_EXPORTS = window.GM_EXPORTS || {};' +
    #13#10 +
    'window.GM_EXPORTS["' + JsonEscape(TabName) + '"] = ' + RowsJson.ToString + ';', Encoding.UTF8);
  
  RebuildManifest(BaseDir);
  Result := ExportedRows;
end;


end.
