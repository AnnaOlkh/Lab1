using Xunit;
using Lab1.MyGrammar;
using System;
using System.Collections.Generic;
using System.IO;

namespace Lab1.UnitTests
{
    public class EvaluatorTests
    {
        private readonly Dictionary<string, string> _cellExpressions;

        public EvaluatorTests()
        {
            _cellExpressions = new Dictionary<string, string>();
        }

        [Fact]
        public void Evaluate_ValidExpression_ReturnsValidResult()
        {
            var evaluator = new Evaluator(_cellExpressions);

            var result = evaluator.Evaluate("(2 + -3 * (9 mod 7)) - min(9, 23, 1 - 1.5)");
            Assert.Equal(-3.5, result);
        }

        [Fact]
        public void Evaluate_DivisionByZero_ThrowsDivideByZeroException()
        {
            var evaluator = new Evaluator(_cellExpressions);

            var exception = Assert.Throws<DivideByZeroException>(() => evaluator.Evaluate("3 / 0"));
            Assert.Equal("Ділення на нуль.", exception.Message);
        }
        [Fact]
        public void Evaluate_CyclicCellReference_ThrowsInvalidOperationException()
        {            _cellExpressions["A1"] = "B1";
            _cellExpressions["B1"] = "A1";
            var evaluator = new Evaluator(_cellExpressions);

            var exception = Assert.Throws<InvalidOperationException>(() => evaluator.Evaluate("A1", "A1"));
            Assert.Equal("Виявлено циклічне посилання на клітинку A1.", exception.Message);
        }

        [Fact]
        public void Evaluate_CellWithMissingReference_ThrowsInvalidDataException()
        {
            var evaluator = new Evaluator(_cellExpressions);
            var exception = Assert.Throws<InvalidDataException>(() => evaluator.Evaluate("F16"));
            Assert.Equal("Не існує клітинки F16", exception.Message);
        }

        [Fact]
        public void Evaluate_EmptyExpression_ThrowsArgumentNullException()
        {
            var evaluator = new Evaluator(_cellExpressions);
            var exception = Assert.Throws<ArgumentNullException>(() => Evaluator.HandleEmptyExpression(""));
            Assert.Equal("Введений вираз пустий. (Parameter 'expr')", exception.Message);
        }

        [Fact]
        public void Evaluate_MinFunction_ReturnsCorrectValue()
        {
            var evaluator = new Evaluator(_cellExpressions);

            var result = evaluator.Evaluate("min(10, 20, -5)");

            Assert.Equal(-5, result);
        }

        [Fact]
        public void Evaluate_MaxFunction_ReturnsCorrectValue()
        {
            var evaluator = new Evaluator(_cellExpressions);
            var result = evaluator.Evaluate("max(10, 20, -5)");
            Assert.Equal(20, result);
        }

        [Fact]
        public void Evaluate_DivideFunction_ReturnsCorrectValue()
        {
            var evaluator = new Evaluator(_cellExpressions);
            var result = evaluator.Evaluate("10 div 3");
            Assert.Equal(3, result);
        }

        [Fact]
        public void Evaluate_ModuloFunction_ReturnsCorrectValue()
        {
            var evaluator = new Evaluator(_cellExpressions);
            var result = evaluator.Evaluate("10 mod 3");
            Assert.Equal(1, result);
        }
    }
}

