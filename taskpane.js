(function(){
  const TEXT = "FERTIG";
  const COLOR = "#00b050";
  const FONT  = "Calibri Light";
  const SIZE  = 10;

  async function insertStyled(){
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const range = sel.insertText(TEXT, "Replace");
      range.font.color = COLOR;
      range.font.name  = FONT;
      range.font.size  = SIZE;
      await context.sync();
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("btn-insert");
    btn && btn.addEventListener("click", () => {
      insertStyled().catch(err => {
        console.error(err);
        alert("Konnte nicht einfügen. Cursor im Dokument platzieren und erneut versuchen.");
      });
    });
  });
})();