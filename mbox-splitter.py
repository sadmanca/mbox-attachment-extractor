# -*- coding: utf-8 -*-

import sys
import os
import mailbox

def split_mbox(filename, split_size_mb):
    try:
        split_size = split_size_mb * 1024 * 1024
        chunk_count = 1
        output = f"{filename.replace('.mbox', f'_{chunk_count}.mbox')}"

        if os.path.exists(output):
            print(f'The file `{filename}` has already been split. Delete chunks to continue.')
            return

        print(f'Splitting `{filename}` into chunks of {split_size_mb} Mb ...\n')

        total_size = 0
        message_count = 0
        output_mailbox = None

        mbox = mailbox.mbox(filename)

        for message in mbox:
            if output_mailbox is None:
                output_mailbox = mailbox.mbox(output, create=True)

            output_mailbox.add(message)
            message_count += 1
            total_size += len(str(message))

            if total_size >= split_size:
                output_mailbox.flush()
                output_mailbox.close()
                print(f'Created file `{output}`, size={total_size / 1024 / 1024:.2f} Mb, messages={message_count}.')
                
                chunk_count += 1
                output = f"{filename.replace('.mbox', f'_{chunk_count}.mbox')}"
                output_mailbox = None
                total_size = 0
                message_count = 0

        if output_mailbox is not None:
            output_mailbox.flush()
            output_mailbox.close()
            print(f'Created file `{output}`, size={total_size / 1024 / 1024:.2f} Mb, messages={message_count}.')

        print('\nDone')

    except Exception as e:
        print(f'Error: {e}')

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python mbox-splitter.py filename.mbox size')
        print('       where size is a positive integer in Mb')
        print()
        print('Example: python mbox-splitter.py inbox_test.mbox 50')
        print('         where inbox_test.mbox size is about 125 Mb')
        print()
        print('Result:')
        print('Created file `inbox_test_1.mbox`, size=43 Mb, messages=35')
        print('Created file `inbox_test_2.mbox`, size=44 Mb, messages=2')
        print('Created file `inbox_test_3.mbox`, size=30 Mb, messages=73')
        print('Done')
        sys.exit(1)

    filename = sys.argv[1]
    if not os.path.exists(filename):
        print(f'File `{filename}` does not exist.')
        sys.exit(1)

    try:
        split_size_mb = int(sys.argv[2])
        if split_size_mb <= 0:
            raise ValueError('Size must be a positive number')
    except ValueError:
        print('Size must be a positive integer')
        sys.exit(1)

    split_mbox(filename, split_size_mb)
