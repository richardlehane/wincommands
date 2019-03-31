// Copyright 2018 State of New South Wales through the State Archives and Records Authority of NSW
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package wincommands

import (
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"
)

// defaults

var (
	tikaInstall        = `C:\apache_tika\tika-app-1.5.jar`
	imageMInstall      = `C:\Program Files\ImageMagick-6.8.8-Q16\convert.exe`
	libreOfficeInstall = `C:\Program Files\LibreOffice 5\program\soffice`
	ffmpegInstall      = `C:\ffmpeg\bin\ffmpeg.exe`
	thumbDimensions    = "1024x1024"
	timeout            = 30 * time.Second

	extract = []string{"java", "-jar", tikaInstall, "-t"}
	thumb   = []string{imageMInstall, "-resize", thumbDimensions, "-flatten", "-quality", "100"}
	pdf     = []string{libreOfficeInstall, "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir"}
	ffmpeg  = []string{ffmpegInstall}
)

// SetFFMpegPath sets your install directory for FFMpeg
func SetFFMpegPath(p string) {
	ffmpegInstall = p
	ffmpeg = []string{ffmpegInstall}
}

// SetTikaPath sets your install directory for Tika
func SetTikaPath(p string) {
	tikaInstall = p
	extract = []string{"java", "-jar", tikaInstall, "-t"}
}

// SetImageMPath sets your install directory for Image Magick
func SetImageMPath(p string) {
	imageMInstall = p
	thumb = []string{imageMInstall, "-resize", thumbDimensions, "-flatten", "-quality", "100"}
}

// SetLibreOPath sets your install directory for Libre Office
func SetLibreOPath(p string) {
	libreOfficeInstall = p
	pdf = []string{libreOfficeInstall, "--headless", "--convert-to", "pdf:writer_pdf_Export", "--outdir"}
}

// SetThumb defines your preferences for thumbnail dimensions (provide x and y values)
func SetThumb(x, y int) {
	thumbDimensions = fmt.Sprintf("%dx%d", x, y)
	thumb = []string{imageMInstall, "-resize", thumbDimensions, "-flatten", "-quality", "100"}
}

// SetTimeout sets a timeout for actions
func SetTimeout(t time.Duration) {
	timeout = t
}

// commands

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

// MakeDir creates a directory, unless one already exists.
// Returns an error and a boolean that indicates whether the directory already exists.
func MakeDir(dir string) (error, bool) {
	err := os.MkdirAll(dir, 0777)
	if os.IsExist(err) {
		return nil, true
	} else if err == nil {
		return nil, false
	}
	return fmt.Errorf("Commands: Error making directory %s, error message: %v", dir, err), false
}

// ExtractText extracts text from input and writes it to the outname in outdir
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

// Thumbnail creates a thumbnail of input in outname in outdir
func Thumbnail(input, outdir, outname string, overwrite bool) error {
	output := filepath.Join(outdir, outname)
	if handleOverwrite(overwrite, output) {
		return nil
	}
	thumbCmd := buildCmd(thumb, input+"[0]", output)
	return timeOutRun(thumbCmd, timeout)
}

func quotePath(path string) string {
	if strings.HasPrefix(path, "\"") {
		return path
	}
	if strings.Contains(path, " ") {
		return "\"" + path + "\""
	}
	return path
}

// FileCopy uses robocopy to copy a file input to outdir
func FileCopy(input, outdir string, overwrite bool) error {
	output := filepath.Join(outdir, filepath.Base(input))
	if handleOverwrite(overwrite, output) {
		return nil
	}
	if err, _ := MakeDir(outdir); err != nil {
		return err
	}
	return fileCopy(input, outdir)
}

// FileCopy log copies a file from input to outdir using xcopy and logs the copy action to the provided log writer
func FileCopyLog(lg io.Writer, input, outdir string, overwrite bool) error {
	output := filepath.Join(outdir, filepath.Base(input))
	if handleOverwrite(overwrite, output) {
		return nil
	}
	if err, _ := MakeDir(outdir); err != nil {
		return err
	}
	cpCmd := copyCmd(input, outdir, true)
	_, err := fmt.Fprintln(lg, strings.Join(cpCmd.Args, " "))
	return err
}

// WordToPdf turns a word doc at input into a PDF file in outdir
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
	_ = timeOutRun(pdfCmd, timeout)
	if _, err := os.Stat(output); err != nil {
		e := os.RemoveAll(outdir) // failed to create, cleanup
		if e != nil {
			return "", fmt.Errorf("Can't create, can't delete: %v, %v", err, e)
		}
		return "", nil
	}
	return output, nil
}

// IsWord tests a PUID against the MS Word formats
func IsWord(puid string) bool {
	switch puid {
	case "fmt/37", "fmt/38", "fmt/39", "fmt/40", "fmt/412", "fmt/523", "fmt/597", "fmt/599", "fmt/609", "fmt/754", "x-fmt/45":
		return true
	}
	return false
}

// IsPDF tests a PUID against the PDF formats
func IsPDF(puid string) bool {
	switch puid {
	case "fmt/14", "fmt/15", "fmt/16", "fmt/17", "fmt/18", "fmt/19", "fmt/20", "fmt/95", "fmt/144", "fmt/145", "fmt/146", "fmt/147", "fmt/148", "fmt/157", "fmt/158", "fmt/276", "fmt/354", "fmt/476", "fmt/477", "fmt/478", "fmt/479", "fmt/480", "fmt/481", "fmt/488", "fmt/489", "fmt/490", "fmt/491", "fmt/492", "fmt/493":
		return true
	}
	return false
}

// IsText tests a PUID against the text formats
func IsText(puid string) bool {
	switch puid {
	case "fmt/14", "fmt/15", "fmt/16", "fmt/17", "fmt/18", "fmt/19", "fmt/20", "fmt/37", "fmt/38", "fmt/39", "fmt/40", "fmt/95", "fmt/144", "fmt/145", "fmt/146", "fmt/147", "fmt/148", "fmt/157", "fmt/158", "fmt/276", "fmt/354", "fmt/412", "fmt/473", "fmt/476", "fmt/477", "fmt/478", "fmt/479", "fmt/480", "fmt/481", "fmt/488", "fmt/489", "fmt/490", "fmt/491", "fmt/492", "fmt/493", "fmt/523", "fmt/597", "fmt/599", "fmt/609", "fmt/754", "x-fmt/45", "x-fmt/111", "x-fmt/273", "x-fmt/274", "x-fmt/275", "x-fmt/276":
		return true
	}
	return false
}
