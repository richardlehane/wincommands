package wincommands

import (
	"os/exec"
)

var (
	cp = []string{"cp"}
)

func copyCmd(input, outdir string, quote bool) *exec.Cmd {
	return buildCmd(cp, input, outdir)
}

func fileCopy(input, outdir string) error {
	cmd := copyCmd(input, outdir, false)
	return cmd.Run()
}
