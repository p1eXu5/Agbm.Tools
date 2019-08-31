using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.Helpers.Monads
{
    public static class MaybyExtensions
    {
        public static Maybe<T> Return<T>( this T instance ) where T : class => new Maybe< T >( instance );

        public static (Maybe<T>, Maybe<T>) BindParallel< T >( this (Maybe<T>, Maybe<T>) monads, Func< Maybe<T>, Maybe<T>> funcA, Func< Maybe< T >, Maybe< T > > funcB )
            where T : class
        {
            return ( funcA( monads.Item1 ), funcB( monads.Item2 ));
        }
    }
}
