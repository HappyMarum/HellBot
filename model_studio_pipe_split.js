/*
  Model Studio ActiveX JScript template
  Goal: split a selected pipe layout area into equal square cells
  so each square can be exported/loaded separately for transport.

  NOTE:
  - Method/property names vary by Model Studio product/version.
  - Replace TODO-marked API calls with your exact object model names.
  - Run with Windows Script Host in an environment where Model Studio COM is installed.
*/

(function () {
  var CFG = {
    // Preferred square side in model units (e.g., mm).
    preferredSquareSide: 1200,

    // Layer/category where helper split lines will be created.
    splitLayerName: "TRUCK_SPLIT_GRID",

    // Whether to create text labels for each square (A1, A2, ...).
    addLabels: true
  };

  function fail(msg) {
    WScript.Echo("ERROR: " + msg);
    WScript.Quit(1);
  }

  function ceilDiv(a, b) {
    return Math.ceil(a / b);
  }

  function toLetter(i) {
    // 0->A, 1->B ... 25->Z, 26->AA...
    var s = "";
    i++;
    while (i > 0) {
      var m = (i - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  // 1) Connect to Model Studio
  var app;
  try {
    // TODO: replace ProgID with your exact Model Studio ProgID.
    app = new ActiveXObject("ModelStudio.Application");
  } catch (e) {
    fail("Cannot create ActiveX object. Check Model Studio ProgID. " + e.message);
  }

  // TODO: adapt these handles to your COM model
  var doc = app.ActiveDocument;
  if (!doc) fail("No active document.");

  var sel = doc.Selection;
  if (!sel || sel.Count < 1) {
    fail("Select one closed contour (pipe layout boundary) first.");
  }

  // 2) Read selected boundary and get bounding box
  // TODO: adjust to your API: e.g. sel.Item(0).GetBoundingBox()
  var boundary = sel.Item(0);
  var bb = boundary.GetBoundingBox(); // expected: {MinX, MinY, MaxX, MaxY}

  var minX = bb.MinX;
  var minY = bb.MinY;
  var maxX = bb.MaxX;
  var maxY = bb.MaxY;

  var width = maxX - minX;
  var height = maxY - minY;

  if (width <= 0 || height <= 0) {
    fail("Boundary has invalid size.");
  }

  // 3) Compute equal-square grid that fully covers boundary box
  var cols = ceilDiv(width, CFG.preferredSquareSide);
  var rows = ceilDiv(height, CFG.preferredSquareSide);

  // Force truly equal square side for both directions.
  var squareSide = Math.max(width / cols, height / rows);

  // Recompute counts with final side (ensures full coverage)
  cols = ceilDiv(width, squareSide);
  rows = ceilDiv(height, squareSide);

  // 4) Prepare output layer/group
  // TODO: replace with actual layer/category API
  var layer = doc.Layers.GetOrCreate(CFG.splitLayerName);

  // 5) Draw vertical and horizontal grid lines
  var c, r;
  for (c = 0; c <= cols; c++) {
    var x = minX + c * squareSide;
    // TODO: clip to boundary if needed using your geometry API.
    doc.ModelSpace.AddLine(x, minY, x, maxY, layer);
  }

  for (r = 0; r <= rows; r++) {
    var y = minY + r * squareSide;
    // TODO: clip to boundary if needed using your geometry API.
    doc.ModelSpace.AddLine(minX, y, maxX, y, layer);
  }

  // 6) Optional labels for logistics (A1, A2...)
  if (CFG.addLabels) {
    for (r = 0; r < rows; r++) {
      for (c = 0; c < cols; c++) {
        var cx = minX + (c + 0.5) * squareSide;
        var cy = minY + (r + 0.5) * squareSide;
        var label = toLetter(r) + (c + 1);
        // TODO: replace text method/signature.
        doc.ModelSpace.AddText(label, cx, cy, squareSide * 0.12, layer);
      }
    }
  }

  // 7) Commit/refresh
  // TODO: adjust to your transaction model.
  doc.Regen();

  WScript.Echo(
    "Done. Grid created: " + rows + " x " + cols +
    ", square side = " + squareSide.toFixed(2)
  );
})();
