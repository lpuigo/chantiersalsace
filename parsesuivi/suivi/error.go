package suivi

type ParsingError struct {
	Msgs []string
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
}

func (e *ParsingError) Add(err error) {
	e.Msgs = append(e.Msgs, err.Error())
}

func (e ParsingError) HasError() bool {
	return len(e.Msgs) > 0
}
