

import os, platform, subprocess
from tempfile import TemporaryDirectory
import zipfile

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolledtext
from tkinter import filedialog

from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName

basedir = os.path.dirname(__file__)

if platform.system() == 'Windows':
    win = True
    LO = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
else:
    win = False
    LO = "libreoffice"

def run_extern(command, *args, cwd = None, xpath = None, feedback = None):
    """Run an external program.
    Pass the command and the arguments as individual strings.
    The command must be either a full path or a command known in the
    run-time environment (PATH).
    Named parameters can be used to set:
     - cwd: working directory. If provided, change to this for the
       operation.
     - xpath: an additional PATH component (prefixed to PATH).
     - feedback: If provided, it should be a function. It will be called
         with each line of output as this becomes available.
    Return a tuple: (return-code, message).
    return-code: 0 -> ok, 1 -> fail, -1 -> command not available.
    If return-code >= 0, return the output as the message.
    If return-code = -1, return a message reporting the command.
    """
    # Note that using the <timeout> parameter will probably not work,
    # at least not as one might expect it to.
    params = {
        'stdout': subprocess.PIPE,
        'stderr': subprocess.STDOUT,
        'universal_newlines':True
    }
    my_env = os.environ.copy()
    if win:
        # Suppress the console
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        params['startupinfo'] = startupinfo
    if xpath:
        # Extend the PATH for the process
        my_env['PATH'] = xpath + os.pathsep + my_env['PATH']
        params['env'] = my_env
    if cwd:
        # Switch working directory for the process
        params['cwd'] = cwd

    cmd = [command] + list(args)
    try:
        if feedback:
            out = []
            with subprocess.Popen(cmd, bufsize=1, **params) as cp:
                for line in cp.stdout:
                    l = line.rstrip()
                    out.append(l)
                    feedback(l)
            msg = '\n'.join(out)

        else:
            cp = subprocess.run(cmd, **params)
            msg = cp.stdout

        return (0 if cp.returncode == 0 else 1, msg)

    except FileNotFoundError:
        return (-1, _COMMANDNOTPOSSIBLE.format(cmd=repr(cmd)))


def libre_office(in_list, pdf_dir):
    """Convert a list of odt (or docx or rtf, etc.) files to pdf-files.
    The input files are provided as a list of absolute paths,
    <pdf_dir> is the absolute path to the output folder.
    """
# Use LibreOffice to convert the input files to pdf-files.
# If using the appimage, the paths MUST be absolute, so I use absolute
# paths "on principle".
# I don't know whether the program blocks until the conversion is complete
# (some versions don't), so it might be good to check that all the
# expected files have been generated (with a timeout in case something
# goes wrong?).
# The old problem that libreoffice wouldn't work headless if another
# instance (e.g. desktop) was running seems to be no longer the case,
# at least on linux.
    def extern_out(line):
        REPORT(line)

    rc, msg = run_extern(LO, '--headless',
            '--convert-to', 'pdf',
            '--outdir', pdf_dir,
            *in_list,
            feedback = extern_out
        )


def merge_pdf(ifile_list, ofile, pad2sided = False):
    """Join the pdf-files in the input list <ifile_list> to produce a
    single pdf-file. The output is returned as a <bytes> object.
    The parameter <pad2sided> allows blank pages to be added
    when input files have an odd number of pages – to ensure that
    double-sided printing works properly.
    """
    writer = PdfWriter()
    for inpfn in ifile_list:
        ipages = PdfReader(inpfn).pages
        if pad2sided and len(ipages) & 1:
            # Make sure we have an even number of pages by copying
            # and blanking the first page
            npage = ipages[0].copy()
            npage.Contents = PdfDict(stream='')
            ipages.append(npage)
        writer.addpages(ipages)
    writer.write(ofile)


def get_input():
    text.delete('1.0', tk.END)

#    files = filedialog.askopenfilenames(
#        parent=root,
##            initialdir='/',
##            initialfile='tmp',
#        filetypes=[
#            ("All files", "*"),
#            ("Word", "*.docx"),
#            ("LibreOffice", "*.odt"),
#            ("RTF", "*.rtf")
#        ]
#    )

    idir = filedialog.askdirectory(parent=root)
    if idir:
        conjoindir(idir)


def get_zip():
    text.delete('1.0', tk.END)

    zfile = filedialog.askopenfilename(
        parent=root,
#            initialdir='/',
#            initialfile='tmp',
        filetypes=[("Zip-Dateien", ".zip*")]
    )
    if zfile:
        with TemporaryDirectory() as zdir:
            # Extract archive
            try:
                with zipfile.ZipFile(zfile, mode="r") as archive:
                    archive.extractall(zdir)
            except zipfile.BadZipFile as error:
                REPORT(f"FEHLER: {error}")
                return

            # Handle files not in a subfolder
            conjoindir(zdir, name=os.path.basename(zfile.rsplit(".", 1)[0]))

            # Handle subfolders
            for d in sorted(os.listdir(zdir)):
                sdir = os.path.join(zdir, d)
                if os.path.isdir(sdir):
                    conjoindir(sdir)


def conjoindir(idir, name=None):
    #print("???", idir)
    files = []
    for f in sorted(os.listdir(idir)):
        try:
            b, e = f.rsplit(".", 1)
        except ValueError:
            continue
        if e in ("odt", "docx", "rtf"):
            files.append(os.path.join(idir, f))

    if files:
        idir = os.path.dirname(files[0])
        with TemporaryDirectory() as odir:

            root.config(cursor="watch")
            text.config(cursor="watch") # Seems to be needed additionally!
            root.update_idletasks()

            libre_office(files, odir)
            root.config(cursor="")
            text.config(cursor="")
            pfiles = []
            REPORT("\n *******************************\n")
            for f in files:
                bpdf = os.path.basename(f).rsplit(".", 1)[0] + ".pdf"
                fpdf = os.path.join(odir, bpdf)
                if os.path.isfile(fpdf):
                    pfiles.append(fpdf)
                else:
                    REPORT(" *** FEHLER, Datei fehlt: {fpdf}")
            if len(files) == len(pfiles):
                sfile = filedialog.asksaveasfilename(
                    parent=root,
                    defaultextension=".pdf",
                    #initialdir=idir,
                    initialfile=(name or os.path.basename(idir)) + ".pdf",
                    filetypes=[("PDF", "*.pdf")]
                )
                if sfile:
                    if not sfile.endswith(".pdf"):
                        sfile += ".pdf"
                    merge_pdf(pfiles, sfile,
                        pad2sided=twosided.instate(['selected'])
                    )
                    REPORT(f" --> {sfile}")


def REPORT(line):
    text.insert(tk.END, line.rstrip() + "\n")
    root.update_idletasks()
    text.yview(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.tk.call('tk_getOpenFile', '-foobarbaz')
    except:
        pass
    try:
        #root.tk.call('set', '::tk::dialog::file::showHiddenBtn', '1')
        root.tk.call('set', '::tk::dialog::file::showHiddenVar', '0')
    except:
        pass

    root.title("Conjoin2pdf")
    root.iconphoto(True, tk.PhotoImage(file=os.path.join(basedir, 'conjoin2pdf.png')))
    #root.geometry('300x300')
    bt = ttk.Button(
        root,
        text="Eingabe von Ordner (mit odt-, docx-, rtf-Dateien)",
        command=get_input
    )
    bt.pack(fill=tk.X, padx=3, pady=3)
    btz = ttk.Button(
        root,
        text="Eingabe von Zip-Datei",
        command=get_zip
    )
    btz.pack(fill=tk.X, padx=3, pady=3)
    twosided = ttk.Checkbutton(root, text="Doppelseitige Ausgabe")
    twosided.pack(fill=tk.X, padx=3, pady=3)
    #twosided.state(['!alternate', '!selected'])
    twosided.state(['!alternate', 'selected'])
    #print("?$?", twosided.state())
    #print("?$?", twosided.instate(['selected']))
    text = scrolledtext.ScrolledText(root)#, state=tk.DISABLED)
    text.bind("<Key>", lambda e: "break")
    text.pack()
    root.update()
    w, h = root.winfo_width(), root.winfo_height()
    #print(f"w={w}, h={h}")
    x = int(root.winfo_screenwidth()/2 - w/2)
    y = int(root.winfo_screenheight()/2 - h/2)
    #x, y = 200, 100
    root.geometry(f"+{x}+{y}")
    root.mainloop()

