package suivi

type Error struct {
	Msgs []string
}

func (e Error) Error() string {
	return e.ToString("\t", "\r\n")
}

func (e Error) ToString(prefix, sep string) string {
	res := ""
	for _, msg := range e.Msgs {
		res += prefix + msg + sep
	}
	return res
}

func (e *Error) Add(err error) {
	e.Msgs = append(e.Msgs, err.Error())
}

func (e Error) HasError() bool {
	return len(e.Msgs) > 0
}
