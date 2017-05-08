import {test} from 'ava';

import { User } from 'excol';

test('should handle getEmail', t => {

  const user: User = new User('email');

  t.is(user.getEmail(), 'email');

});
