package suivi

type ParsingError struct {
	Msgs  []string
	Fatal bool
}

func (e ParsingError) Error() string {
	return e.ToString("\t", "\r\n")
}

func (e ParsingError) ToString(prefix, sep string) string {
	res := ""
	for _, msg := range e.Msgs {
		res += prefix + msg + sep
	}
	return res
}

func (e *ParsingError) Append(perr ParsingError) {
	e.Msgs = append(e.Msgs, perr.Msgs...)
	e.Fatal = e.Fatal || perr.Fatal
}

func (e *ParsingError) Add(err error, fatal bool) {
	e.Msgs = append(e.Msgs, err.Error())
	e.Fatal = e.Fatal || fatal
}

func (e ParsingError) HasError() bool {
	return len(e.Msgs) > 0
}

func (e ParsingError) IsFatal() bool {
	return e.Fatal
}

func (e ParsingError) SetFatal(fatal bool) ParsingError {
	e.Fatal = e.Fatal || fatal
	return e
}
