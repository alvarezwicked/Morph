using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TangerineWin.Models;

public sealed record FileAction(
    string Id,
    string DisplayName
);
