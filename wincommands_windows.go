package wincommands

import (
	"fmt"
	"os/exec"
)

var (
	rcp = []string{"robocopy"} //"/COPY:DATSO" - fails over NTFS (security), /COPYALL fails don't have manage user auditing right "/LOG:c:\\Users\\richardl\\Desktop\\log.txt"
	xcp = []string{"xcopy"}
)

func fileCopy(input, outdir) error {
	return runRobo(input, outdir)
}

func copyCmd(input, outdir string, quote bool) *exec.Cmd {
	return xcopy(input, outdir, quote)
}

func runRobo(input, outdir string) error {
	cpCmd := robo(input, outdir, false)
	err := cpCmd.Run()
	if err == nil {
		return fmt.Errorf("Commands: Error copying %s to %s with command %s, error message: No errors occurred and no files were copied", input, outdir, strings.Join(append([]string{cpCmd.Path}, cpCmd.Args...), " "))
	}
	if err.Error() == "exit status 1" {
		return nil
	}
	return fmt.Errorf("Commands: Error copying %s to %s with command %s, error message: %v", input, outdir, strings.Join(append([]string{cpCmd.Path}, cpCmd.Args...), " "), err)
}

func robo(input, outdir string, quote bool) *exec.Cmd {
	dir, fn := filepath.Split(input)
	if len(dir) > 0 {
		dir = dir[:len(dir)-1]
	}
	if quote {
		dir, outdir, fn = quotePath(dir), quotePath(outdir), quotePath(fn)
	}
	return buildCmd(rcp, dir, outdir, fn)
}

func xcopy(input, outdir string, quote bool) *exec.Cmd {
	if quote {
		input, outdir = quotePath(input), quotePath(outdir)
	}
	return buildCmd(xcp, input, outdir)
}

func runXcopy(input, outdir string) error {
	cpCmd := xcopy(input, outdir, false)
	err := cpCmd.Run()
	if err == nil {
		return nil
	}
	return fmt.Errorf("Commands: Error copying %s to %s with command %s, error message: %v", input, outdir, strings.Join(append([]string{cpCmd.Path}, cpCmd.Args...), " "), err)
}
