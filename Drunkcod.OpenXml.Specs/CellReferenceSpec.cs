using Cone;
using System;

namespace Drunkcod.OpenXml.Specs
{
    [Describe(typeof(CellReference))]
    public class CellReferenceSpec
    {
        public void empty_is_empty_string() => 
            Check.That(() => new CellReference().ToString() == string.Empty);

        public void can_create_invalid_referencet() {
            Check.Exception<InvalidOperationException>(() => new CellReference(0, 0));
            Check.Exception<InvalidOperationException>(() => new CellReference(-1, 1));
            Check.Exception<InvalidOperationException>(() => new CellReference(1, -1));
        }
    }
}
