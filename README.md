# VirtualBox Remote Snapshots

An attempt at creating a new external snapshot system for VirtualBox, leveraging the following existing projects:

* [`borg` backup](https://github.com/borgbackup/borg)
* [`rsync`](https://rsync.samba.org/)
* [`VSS`](https://msdn.microsoft.com/en-us/library/windows/desktop/bb968832(v=vs.85).aspx)

## Motivation

The native VirtualBox snapshot system has the following properties:

#### 1. Snapshot space requirements

For a quickly changing virtual machine, the required space for virtual machine snapshots in VirtualBox can be quite large. 
This is at odds with the small amounts of avaliable solid state storage on inexpensive laptops.
To accomodate these storage needs, snapsnots are kept for short periods of time, so they won't have time to increase in size, and are taken infrequently, leading to poor snapshot coverage.

In this system, snapshots are kept on inexpensive bulk storage on a separate machine, which can host an order of magnitude more snapshots without consuming precious local resources.

#### 2. Snapshot crash consistency

If a VirtualBox host goes down during a snapshot operation, the guest's disk may be left in an inconsistent state, rendering it unusable.

In this system, snapshots operations are incremental and resumable. 
If the host goes down, a snapshot restore can simply continue where it left off.

#### 3. Snapshot performance implications

The CoW magic that makes VirtualBox snapshots effectively instantaneous incurs a performance penalty on future disk I/O in the guest. 
The performance penalty increases with the number of snapshots.

In this system, there is generally no runtime penalty as long as a snaphot is not _currently being taken_. 
If no snapshot is being taken, this snapshot system does not even need to be running. 

## Architecture 

Snapshots are stored remotely in a `borg` repository. 
Since `borg` only trasmits changed blocks which are then compressed, a full snapshot can often be taken in under a minute, even over a slow WiFi connection. 

Snapshots can be taken while the guest is online. 
Transient VSS snapshots are used to ensure that the guests's disk file is crash-consistent. 
Support for fully consistent online snapshots is planned.

Snapshot restores are incremental and resumable using a combination of `rsync` and `borg mount`. 
Since only modified blocks are transmitted, a recent snapshot can often be restored in under a minute, even over a slow WiFi connection.

Since `borg` is a forever-incremental system, snapshot restores run in constant time with respect to the total number of snapshots.

Snapshots can be pruned with no interaction from the host at all, online or offline.

# Disclaimer 

This system is still in a very early state. Lots of important functionality is missing, and many design decisions still need to be made. The code is TBH pretty horrifying. 
