package wincommands

import (
	"fmt"
	"io/ioutil"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"
)

// install locations
const (
	tikaInstall        = `C:\apache_tika\tika-app-1.5.jar`
	imageMInstall      = `C:\Program Files\ImageMagick-6.8.8-Q16\convert.exe`
	libreOfficeInstall = `C:\Program Files\LibreOffice 5\program\soffice`
	thumbDimensions    = "1024x1024"
)

// commands
var (
	extract = []string{"java", "-jar", tikaInstall, "-t"}
	cp      = []string{"cmd", "/c", "copy"}
	thumb   = []string{imageMInstall, "-resize", thumbDimensions, "-flatten", "-quality", "100"}
	pdf     = []string{libreOfficeInstall, "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir"}
)

func buildCmd(template []string, custom ...string) *exec.Cmd {
	cmd := make([]string, len(template)+len(custom))
	copy(cmd, template)
	copy(cmd[len(template):], custom)
	return exec.Command(cmd[0], cmd[1:]...)
}

func timeOutRun(cmd *exec.Cmd, dur time.Duration) error {
	err := cmd.Start()
	if err != nil {
		return err
	}
	timer := time.AfterFunc(dur, func() {
		e := cmd.Process.Kill()
		if e != nil {
			panic(e)
		}
	})
	err = cmd.Wait()
	timer.Stop()
	return err
}

func handleOverwrite(overwrite bool, output string) bool {
	if !overwrite {
		if _, err := os.Stat(output); err == nil {
			return true
		}
	}
	return false
}

func MakeDir(dir string) (error, bool) {
	err := os.MkdirAll(dir, os.ModePerm)
	if os.IsExist(err) {
		return nil, true
	} else if err == nil {
		return nil, false
	}
	return fmt.Errorf("Commands: Error making directory %s, error message: %v", dir, err), false
}

func ExtractText(input, outdir, outname string, overwrite bool) error {
	output := filepath.Join(outdir, outname)
	if handleOverwrite(overwrite, output) {
		return nil
	}
	if err, _ := MakeDir(outdir); err != nil {
		return err
	}
	tikaCmd := buildCmd(extract, input)
	txt, err := tikaCmd.Output()
	if err != nil {
		return nil // :( no text
		// return fmt.Errorf("Commands: Error making text from %s, error message: %v", input, err)
	}
	err = ioutil.WriteFile(output, txt, os.ModePerm)
	if err != nil {
		return fmt.Errorf("Commands: Error writing to %s, error message: %v", output, err)
	}
	return nil
}

func Thumbnail(input, outdir, outname string, overwrite bool) error {
	output := filepath.Join(outdir, outname)
	if handleOverwrite(overwrite, output) {
		return nil
	}
	thumbCmd := buildCmd(thumb, input+"[0]", output)
	timeOutRun(thumbCmd, 30*time.Second)
	return nil
}

func FileCopy(input, outdir string, overwrite bool) error {
	output := filepath.Join(outdir, filepath.Base(input))
	if handleOverwrite(overwrite, output) {
		return nil
	}
	if err, _ := MakeDir(outdir); err != nil {
		return err
	}
	cpCmd := buildCmd(cp, input, outdir)
	if err := cpCmd.Run(); err != nil {
		return fmt.Errorf("Commands: Error copying %s to %s, error message: %v", input, outdir, err)
	}
	return nil
}

func WordToPdf(input, outdir string, overwrite bool) (string, error) {
	var output string
	switch filepath.Ext(input) {
	case ".doc", ".DOC", ".docx", ".DOCX", ".dotx", ".DOTX", ".docm", ".DOCM":
		output = filepath.Join(outdir, strings.TrimSuffix(filepath.Base(input), filepath.Ext(input))+".pdf")
	default:
		output = filepath.Join(outdir, filepath.Base(input)+".pdf")
	}

	if handleOverwrite(overwrite, output) {
		return output, nil
	}
	if err, _ := MakeDir(outdir); err != nil {
		return "", err
	}
	pdfCmd := buildCmd(pdf, outdir, input)
	timeOutRun(pdfCmd, 30*time.Second)
	if _, err := os.Stat(output); err != nil {
		e := os.RemoveAll(outdir) // failed to create, cleanup
		if e != nil {
			return "", fmt.Errorf("Can't create, can't delete: %v, %v", err, e)
		}
		return "", nil
	}
	return output, nil
}

func IsWord(puid string) bool {
	switch puid {
	case "fmt/37", "fmt/38", "fmt/39", "fmt/40", "fmt/412", "fmt/523", "fmt/597", "fmt/599", "fmt/609", "fmt/754", "x-fmt/45":
		return true
	}
	return false
}

func IsPDF(puid string) bool {
	switch puid {
	case "fmt/14", "fmt/15", "fmt/16", "fmt/17", "fmt/18", "fmt/19", "fmt/20", "fmt/95", "fmt/144", "fmt/145", "fmt/146", "fmt/147", "fmt/148", "fmt/157", "fmt/158", "fmt/276", "fmt/354", "fmt/476", "fmt/477", "fmt/478", "fmt/479", "fmt/480", "fmt/481", "fmt/488", "fmt/489", "fmt/490", "fmt/491", "fmt/492", "fmt/493":
		return true
	}
	return false
}

func IsText(puid string) bool {
	switch puid {
	case "fmt/14", "fmt/15", "fmt/16", "fmt/17", "fmt/18", "fmt/19", "fmt/20", "fmt/37", "fmt/38", "fmt/39", "fmt/40", "fmt/95", "fmt/144", "fmt/145", "fmt/146", "fmt/147", "fmt/148", "fmt/157", "fmt/158", "fmt/276", "fmt/354", "fmt/412", "fmt/473", "fmt/476", "fmt/477", "fmt/478", "fmt/479", "fmt/480", "fmt/481", "fmt/488", "fmt/489", "fmt/490", "fmt/491", "fmt/492", "fmt/493", "fmt/523", "fmt/597", "fmt/599", "fmt/609", "fmt/754", "x-fmt/45", "x-fmt/111", "x-fmt/273", "x-fmt/274", "x-fmt/275", "x-fmt/276":
		return true
	}
	return false
}